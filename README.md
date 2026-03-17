# Dex2

Dex2 is a Chrome extension plus FastAPI backend. The extension captures screenshots of visited pages, sends them to the backend for text extraction and embedding, and a chat UI queries the backend with RAG (retrieval-augmented generation). The backend uses hybrid retrieval (vector + BM25) and a tool-calling agent: the LLM (OpenAI GPT-4o-mini) chooses among tools (open tab, send email, edit slides) or answers in text. The API returns a single action per request (chat only, open tab, send email, or edit slides), and the extension executes the corresponding action.

## Architecture

![Dex2 architecture diagram](new-diagram.png)

### Demo

[Watch the demo](https://www.youtube.com/watch?v=fhYGrfLmNnQ)

---

## Project structure

- **backend/** – FastAPI app: embedding pipeline, hybrid retrieval, chat with tool-calling agent (open_tab, send_email, edit_slides), Google OAuth and Sheets/Docs/Slides integration.
- **frontend/** – Chrome extension (Manifest V3): popup chat UI (React), background service worker (screenshots, Google token storage), and tools (open tab, open Gmail compose, edit slides).

---

## Backend

### Hybrid retrieval

Retrieval combines two strategies over the same document set (user-scoped chunks in MongoDB):

1. **Vector (semantic) search** – The query is embedded with Gemini (`models/gemini-embedding-001`). Each stored chunk has an embedding; cosine similarity between query and chunk gives a score. This finds content that *means* the same thing (e.g. "that database thing" matching a chunk about Postgres).

2. **BM25 (keyword) search** – The query and chunk texts are tokenized (lowercase, split on non-alphanumeric). BM25Okapi scores chunks by exact token overlap. This finds names, URLs, IDs, and dates that vectors can miss (e.g. "did I visit docs.google.com/spreadsheets/..." matching a chunk containing that URL).

Scores are merged with configurable weights:

- `finalScore = (vector_weight * vectorScore) + (text_weight * textScore)`
- Default: 70% vector, 30% text (`RETRIEVAL_VECTOR_WEIGHT=0.7`, `RETRIEVAL_TEXT_WEIGHT=0.3`).
- Vector scores are clamped to [0, 1]; BM25 scores are normalized by the max over the result set so both contribute in a similar range.
- Chunks with `finalScore` below a threshold are dropped (default `RETRIEVAL_MIN_SCORE=0.35`). Remaining chunks are sorted by `finalScore` and the top `k` are returned.

All three values are configurable via environment variables in `langchain_agent.py` (or by passing `vector_weight`, `text_weight`, `min_score` into `retrieve_documents`).

### Tool-calling agent and context

Chat uses RAG plus a single-round tool-calling flow:

1. **Context** – The prompt is built from: (a) retrieved chunks from hybrid search (top-k, default k=4, user-scoped by Supabase token), (b) the user message, (c) the current browser tab URL (when provided), (d) conversation history, and (e) current time. If the frontend sends `current_slide_screenshot`, it is not added to the chat prompt but is passed into the slides orchestrator when the model calls **edit_slides** (for vision-based style). This acts as a context composer so the LLM can choose tools or answer using the right information.

2. **Tools** – The LLM (OpenAI GPT-4o-mini via `langchain-openai`) is given three tools via `bind_tools`:
   - **open_tab** – When the user wants to open a URL or search; parameters: `url` (required), optional `message`. The backend returns `action: "open_tab"` and `msg` containing the message and URL so the frontend can open the tab.
   - **send_email** – When the user wants to compose/send an email; parameters: `email_to`, `email_subject`, `email_body`. The backend builds a Gmail compose URL and returns `action: "send_email"`, `msg`, and `email_url`.
   - **edit_slides** – When the user is on a Google Slides tab (URL contains `docs.google.com/presentation`) and asks to modify or query the presentation; no parameters. The backend calls the slides orchestrator with the current tab URL, user message, Google access token, and optional slide screenshot. When a screenshot is sent and the user asks to **add content** (e.g. a new section or column), Gemini generates the create_shape instructions from the image and layout context (FREE SPACE, element positions); for other operations or when no screenshot is sent, the GPT executor runs. The screenshot is also used for vision-based style so new content matches the slide’s font and colors. Returns `action: "edit_slides"` and `msg` with the result.

3. **Single round** – One LLM invocation per request. If the model returns tool calls, the backend executes only the first tool (since the frontend supports one action per response), maps it to the response enum, and returns. If the model responds with text only, the backend returns `action: "chat_only"` and the model’s message.

4. **Respond** – The API returns `ChatResponse`: `action` (`chat_only`, `open_tab`, `send_email`, or `edit_slides`), `msg`, and optionally `email_url`. The frontend displays `msg` and, depending on `action`, opens a tab from `msg`, opens Gmail compose with `email_url`, handles edit_slides (e.g. show result or open Slides), or shows the message only.

### Google Slides editing

When the user is on a Google Slides tab and asks to modify or query the presentation, the main agent can call the **edit_slides** tool. The backend then runs the slides pipeline in `backend/slides/`:

1. **Entry** – `handle_edit_slides(current_tab_url, user_message, access_token, slide_screenshot=None)` requires a valid Google access token (from Connect Google) and a tab URL that is a Google Slides presentation with a slide fragment (e.g. `#slide=id.xxx`). If the token or URL is missing or invalid, it returns a short error message instead of calling the API. The frontend may send an optional **current_slide_screenshot** (base64 image) when the tab is a Google Slides URL; see “Vision-based slide style” below.

2. **Fetch** – The presentation ID and current slide ID are parsed from the URL. The Google Slides API is used to fetch the full presentation (layout, page size, all slides and elements). The current slide’s elements and free space are computed for context.

3. **Router** – An LLM (same GPT-4o-mini) classifies the user request into one operation type: `answer_question` (Q&A about the deck), `edit_layout` (move, resize, align, center, make symmetrical), `create_content` (add shapes, text boxes, lines on the current slide), `create_slide` (add a new blank slide), or `edit_text` (change text content, font, size, color). The router uses a short context (slide count, title, current slide index, dimensions, element count, free space) so it stays fast.

4. **Executor** – For most operations, a second LLM call (GPT-4o-mini) runs the operation-specific prompt with full presentation and current-slide context (`full_desc`: element list, FREE SPACE gaps, positions). The model outputs structured instructions (create/update/delete elements with positions, sizes, text, style). For `answer_question`, the executor returns a direct answer and no instructions. **Exception:** for **create_content** when a **screenshot** is present, **Gemini** (not GPT) generates the instructions directly from the image, user message, and the same layout context (see “Create content with screenshot” below).

5. **Style** – Before applying instructions, the backend needs primary font, text color, shape fill, and border color so new/edited content matches the deck. **Vision-based style** (Gemini) is used only when adding new elements (**create_content**, **create_slide**) and a screenshot is present; **edit_layout** and **edit_text** use API-based style only. When vision is used, the backend skips the Slides API style fetch and uses Gemini’s extracted style (see “Vision-based slide style” below).

6. **Apply** – For non–answer_question operations, `apply_instructions` translates the instructions into Google Slides API batch updates (create shapes, insert text, update transforms and style). For `create_slide`, the backend first creates a blank slide at the requested index, then applies the generated content to that slide. Success and error messages are returned as plain text.

7. **Response** – The orchestrator returns a string (e.g. "Done! I added 2 elements and updated 1 element. Refresh your Slides tab to see the changes." or an error). The main chat endpoint puts this in `msg` and sets `action: "edit_slides"` so the frontend can show it.

**Create content with screenshot** – When the user asks to add content (e.g. “add a new section next to 2”) and the request includes `current_slide_screenshot`, the backend uses **Gemini** to generate the create_shape instructions directly from the image and the user message. Gemini receives the same **layout context** as the GPT executor (`full_desc`): current slide dimensions, **FREE SPACE** gaps (y ranges where new elements can go), and the list of existing elements with positions and sizes. So Gemini can place the new column or block in the correct empty region and match the visual style it sees in the screenshot. If Gemini returns no instructions, the backend falls back to the GPT executor with a style blurb. Instructions from Gemini are normalized with “fill missing only” so the model’s style choices are preserved.

**Vision-based slide style** – When a screenshot is present and the operation is **create_content** or **create_slide**, the backend uses **Gemini** (`gemini-2.5-flash`) in `backend/slides/vision_style.py` to extract style from the image: primary text color, font, background fill, border color, and (when the slide has a column/card layout) section fill, section border, inner text box fill/border, and title color. These values are used to normalize create_shape instructions; for Gemini-generated content, only missing style fields are filled so the vision model’s choices are kept. **edit_layout** and **edit_text** do not use vision; they use `get_presentation_style_values` (Slides API) only.

Slides code lives under `backend/slides/`: `orchestrator.py` (entry and flow), `router.py` (routing), `executors.py` (operation prompts and GPT executor calls), `actions.py` (apply_instructions and API calls), `api.py` (Slides API helpers and batch updates), `context.py` (presentation context, FREE SPACE gaps, API-based style), `layout.py` (role-based layout, normalize_instructions_style), `vision_style.py` (Gemini: style extraction and create_content instruction generation from image + layout context).

### API endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/` | Root; returns status. |
| GET | `/health` | Health check. |
| GET | `/api/items/{item_id}` | Example item (optional `q`). |
| POST | `/chat` | Chat with RAG and tool-calling. Body: `{ "message": string, "conversation_history": [{ "role", "content" }], "current_tab_url": string \| null, "current_slide_screenshot": string \| null }`. Optional `current_slide_screenshot` is base64 image data (when the user is on a Google Slides tab); used for vision-based style and for Gemini-generated create_content instructions (placement + style from image and layout context). Headers: `Authorization: Bearer <supabase_jwt>`, optional `X-Google-Access-Token` (for edit_slides). Returns: `{ "action": "chat_only" \| "open_tab" \| "send_email" \| "edit_slides", "msg": string, "email_url": string \| null }`. |
| POST | `/api/embed-screenshot/` | Accept screenshot and URL; extract text, then enqueue embedding. Body: `ScreenshotRequest` (source_url, captured_at, title?, screenshot_data). Headers: `Authorization: Bearer <supabase_jwt>`, optional `X-Google-Access-Token` for Google Sheets/Docs. Returns 200 with status; 400 if text extraction fails. URLs under `accounts.google.com` are skipped (not embedded). |
| POST | `/api/google-auth/code` | Exchange OAuth code for tokens. Body: `{ "code", "redirect_uri" }`. Returns `{ "access_token", "refresh_token", "expires_in" }`. |
| POST | `/api/google-auth/refresh` | Refresh access token. Body: `{ "refresh_token" }`. Returns `{ "access_token", "expires_in" }`. |

### Backend setup

- Python 3, venv recommended. Install: `pip install -r requirements.txt`.
- Environment (e.g. `backend/.env`): `GOOGLE_API_KEY` (Gemini: embeddings and vision-based slide style), `OPENAI_API_KEY`, `MONGO_USERNAME`, `MONGO_PASSWORD`; for Google OAuth and Sheets/Docs: `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`. Optional retrieval tuning: `RETRIEVAL_VECTOR_WEIGHT`, `RETRIEVAL_TEXT_WEIGHT`, `RETRIEVAL_MIN_SCORE`.
- Run: `uvicorn main:app --reload` (or `python main.py`). API: http://localhost:8000; docs: http://localhost:8000/docs.

---

## Frontend (Chrome extension)

### Structure

- **Popup** – React app (Vite) loaded when the user clicks the extension icon. Shows login/signup (Supabase) and, when authenticated, the chat UI. "Connect Google" runs OAuth in the popup; the authorization code is sent to the background script, which exchanges it with the backend and stores tokens in `chrome.storage.local` so they persist after the popup closes.
- **Background service worker** – Listens for tab activation and tab updates; captures the visible tab with `chrome.tabs.captureVisibleTab`, then POSTs the screenshot and URL to `/api/embed-screenshot/`. Attaches Supabase JWT and, when available, Google access token (from storage, refreshed if needed). Handles the `GOOGLE_AUTH_SAVE` message from the popup to perform the code exchange and write tokens to storage. When the popup sends `CAPTURE_TAB` with a `windowId`, it captures that window’s visible tab and returns the screenshot as base64 (used when the user is on a Google Slides tab so the chat request can include `current_slide_screenshot` for vision-based style).
- **Tools** – `openTabFromMessage(content)`: parses the first URL from the assistant message and opens it in a new tab (used for `open_tab`). `openEmailCompose(emailUrl)`: opens the Gmail compose URL in a new tab (used for `send_email`).

### Action handling in the popup

After each chat response, the frontend reads `action`, `msg`, and optionally `email_url`. If `action === "open_tab"`, it calls `openTabFromMessage(msg)`. If `action === "send_email"` and `email_url` is present, it calls `openEmailCompose(email_url)`. If `action === "edit_slides"`, it shows the slides result in `msg` (and may open or focus the Slides tab as needed). Otherwise it only shows the message (chat only).

### Frontend setup

- Node 18+. Install: `npm install`. Env: `frontend/.env` with `VITE_SUPABASE_URL`, `VITE_SUPABASE_ANON_KEY`, and `VITE_GOOGLE_CLIENT_ID` (for Connect Google). Build: `npm run build`. Load the **built** extension from `frontend/dist` in Chrome (chrome://extensions, "Load unpacked", select `dist`). The backend must be running (e.g. http://localhost:8000) and the extension's API_URL in the background script must match.

### Google Sheets, Docs, and Slides

To embed content from Google Sheets or Docs, the user clicks "Connect Google" in the popup and completes OAuth (Web application client; redirect URI `https://<extension-id>.chromiumapp.org/`). The backend then uses the user's access token with the Sheets and Docs APIs to extract text when the screenshot URL is a Google Sheets or Docs link. Without Connect Google, those URLs return 401 and are not embedded.

The same Google token is sent with chat requests (header `X-Google-Access-Token`) so that when the user is on a Google Slides tab and asks to edit the presentation, the **edit_slides** tool can call the Google Slides API (fetch presentation, route the request, run the executor or Gemini path, and apply batch updates). When the user sends a message from a Google Slides tab, the popup asks the background to capture that tab (`CAPTURE_TAB` with `windowId`); the base64 screenshot is sent as `current_slide_screenshot`. For “add content” requests (e.g. add a section or column), Gemini uses the screenshot plus layout context (FREE SPACE, element positions) to generate create_shape instructions so new content is placed in the right empty space and matches the slide’s style. The screenshot is also used to extract style (font, text color, fill, border) for create_content and create_slide. Slides editing requires the tab URL to point to a presentation with a slide fragment (e.g. `#slide=id.xxx`). Google Slides presentation URLs are not used for screenshot text extraction (they typically return 401); only the edit_slides flow uses the Slides API.

---

## Running the full stack

1. Start the backend from `backend/`: `uvicorn main:app --reload` (or `python main.py`).
2. Build the frontend from `frontend/`: `npm run build`.
3. In Chrome, load the unpacked extension from `frontend/dist`.
4. Open the extension popup, sign in (Supabase), optionally Connect Google, and use the chat. Visiting pages will capture screenshots and send them to the backend for embedding; chat uses hybrid retrieval and the tool-calling agent to return one action per request (chat only, open tab, send email, or edit slides) that the extension can execute.
