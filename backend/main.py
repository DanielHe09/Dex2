import asyncio
import os
from fastapi import FastAPI, HTTPException, Header, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from typing import Optional
from langchain_text_splitters import RecursiveCharacterTextSplitter
import google.generativeai as genai
from mongo_client import get_embeddings_collection
from langchain_agent import retrieve_documents, llm
from datetime import datetime
import uuid
from dotenv import load_dotenv
from pydantic import BaseModel
from typing import Optional, List
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage
import base64
import httpx
from bs4 import BeautifulSoup
from enum import Enum
import json
import re
from urllib.parse import urlencode
from slides_module import handle_edit_slides

load_dotenv()

app = FastAPI(title="FastAPI Server", version="1.0.0")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
async def root():
    """Root endpoint"""
    return {"message": "Welcome to FastAPI Server", "status": "running"}


@app.get("/health")
async def health():
    """Health check endpoint"""
    return {"status": "healthy"}


@app.get("/api/items/{item_id}")
async def read_item(item_id: int, q: str = None):
    """Example endpoint with path and query parameters"""
    return {"item_id": item_id, "q": q}


class Message(BaseModel):
    """Message model for conversation history"""
    role: str  # "user" or "assistant"
    content: str


class ChatRequest(BaseModel):
    """Request model for chat endpoint"""
    message: str
    conversation_history: Optional[List[Message]] = []
    current_tab_url: Optional[str] = None


class ActionType(str, Enum):
    """Action types for chat responses"""
    CHAT_ONLY = "chat_only"
    OPEN_TAB = "open_tab"
    SEND_EMAIL = "send_email"
    EDIT_SLIDES = "edit_slides"


class ChatResponse(BaseModel):
    """Response model for chat endpoint"""
    action: ActionType
    msg: str
    email_url: Optional[str] = None  # Gmail compose URL when action is send_email


class ScreenshotRequest(BaseModel):
    """Request model for screenshot embedding"""
    source_url: str
    captured_at: str  # ISO string
    title: Optional[str] = None
    screenshot_data: str  # Base64 encoded image data (optional, stored but not used for text extraction)


class GoogleAuthCodeRequest(BaseModel):
    """Exchange OAuth authorization code for tokens"""
    code: str
    redirect_uri: str


class GoogleAuthRefreshRequest(BaseModel):
    """Refresh Google access token"""
    refresh_token: str


# ----- Google Sheets/Docs (user OAuth token) -----
def _parse_google_sheets_url(url: str) -> Optional[tuple[str, int]]:
    """Return (spreadsheet_id, gid) or None. gid defaults to 0."""
    if not url or "docs.google.com/spreadsheets" not in url:
        return None
    try:
        # .../d/SPREADSHEET_ID/edit?gid=0#gid=0
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
        if not match:
            return None
        spreadsheet_id = match.group(1)
        gid_match = re.search(r"[?&#]gid=(\d+)", url)
        gid = int(gid_match.group(1)) if gid_match else 0
        return (spreadsheet_id, gid)
    except Exception:
        return None


def _parse_google_docs_url(url: str) -> Optional[str]:
    """Return document_id or None."""
    if not url or "docs.google.com/document" not in url:
        return None
    try:
        match = re.search(r"/document/d/([a-zA-Z0-9_-]+)", url)
        return match.group(1) if match else None
    except Exception:
        return None


def _extract_text_from_google_sheets(url: str, access_token: str) -> str:
    """Use Sheets API with user token. Returns plain text of cell values."""
    parsed = _parse_google_sheets_url(url)
    if not parsed:
        return ""
    spreadsheet_id, gid = parsed
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        # Get spreadsheet metadata to find sheet name for gid
        meta = httpx.get(
            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}",
            headers=headers,
            params={"fields": "sheets(properties(title,sheetId))"},
            timeout=15.0,
        )
        meta.raise_for_status()
        data = meta.json()
        sheets = data.get("sheets") or []
        sheet_title = None
        for s in sheets:
            props = s.get("properties") or {}
            if props.get("sheetId") == gid:
                sheet_title = (props.get("title") or "Sheet1").strip()
                break
        if not sheet_title:
            sheet_title = (sheets[0].get("properties") or {}).get("title") or "Sheet1"
        # Fetch values (whole sheet)
        range_param = f"'{sheet_title}'!A:ZZ"
        r = httpx.get(
            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{range_param}",
            headers=headers,
            timeout=15.0,
        )
        r.raise_for_status()
        values = r.json().get("values") or []
        lines = []
        for row in values:
            lines.append("\t".join(str(c) for c in row))
        return "\n".join(lines)
    except Exception as e:
        print(f"Error extracting from Google Sheets {url}: {e}")
        return ""


def _extract_text_from_google_docs(url: str, access_token: str) -> str:
    """Use Docs API with user token. Returns plain text."""
    doc_id = _parse_google_docs_url(url)
    if not doc_id:
        return ""
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        r = httpx.get(
            f"https://docs.googleapis.com/v1/documents/{doc_id}",
            headers=headers,
            timeout=15.0,
        )
        r.raise_for_status()
        doc = r.json()
        body = doc.get("body") or {}
        content = body.get("content") or []
        parts = []

        def extract_text_from_element(el: dict) -> None:
            if "paragraph" in el:
                for elem in (el["paragraph"].get("elements") or []):
                    run = elem.get("textRun")
                    if run:
                        parts.append((run.get("content") or "").strip())
            if "table" in el:
                for row in (el["table"].get("tableRows") or []):
                    for cell in (row.get("tableCells") or []):
                        for c in (cell.get("content") or []):
                            extract_text_from_element(c)
            if "tableOfContents" in el:
                toc = el["tableOfContents"]
                for c in (toc.get("content") or []):
                    extract_text_from_element(c)

        for el in content:
            extract_text_from_element(el)
        return " ".join(p for p in parts if p)
    except Exception as e:
        print(f"Error extracting from Google Docs {url}: {e}")
        return ""


def extract_text_from_url(url: str, google_access_token: Optional[str] = None) -> str:
    """
    Extract text content from a webpage URL.
    For Google Sheets/Docs URLs, uses the given user OAuth token and Sheets/Docs API if provided.
    Otherwise fetches the page and parses HTML.
    """
    if google_access_token and (url or "").strip():
        if _parse_google_sheets_url(url):
            text = _extract_text_from_google_sheets(url, google_access_token)
            if text:
                return text
        if _parse_google_docs_url(url):
            text = _extract_text_from_google_docs(url, google_access_token)
            if text:
                return text
    try:
        response = httpx.get(url, timeout=10.0, follow_redirects=True)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        for script in soup(["script", "style"]):
            script.decompose()
        text = soup.get_text(separator=' ', strip=True)
        lines = (line.strip() for line in text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        text = ' '.join(chunk for chunk in chunks if chunk)
        return text if text else ""
    except Exception as e:
        print(f"Error extracting text from URL {url}: {str(e)}")
        return ""


def _run_embedding_sync(
    text: str,
    supabase_token: Optional[str],
    source_url: str,
    captured_at: str,
    title: Optional[str],
    screenshot_data: str,
) -> None:
    """
    Synchronous embedding work (chunking, API calls, MongoDB insert).
    Run in a thread so the main worker stays free for chat and other requests.
    """
    try:
        genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
        text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=1000,
            chunk_overlap=200,
            length_function=len,
        )
        chunks = text_splitter.split_text(text)
        embeddings_collection = get_embeddings_collection()
        document_id = str(uuid.uuid4())
        documents_to_insert = []
        for i, chunk in enumerate(chunks):
            result = genai.embed_content(
                model="models/gemini-embedding-001",
                content=chunk,
                task_type="retrieval_document"
            )
            embedding = result['embedding']
            doc = {
                "supabase_token": supabase_token,
                "document_id": document_id,
                "source_type": "web_screenshot",
                "source_url": source_url,
                "captured_at": captured_at,
                "title": title,
                "filename": source_url,
                "chunk_index": i,
                "text": chunk,
                "embedding": embedding,
                "created_at": datetime.utcnow(),
                "metadata": {
                    "total_chunks": len(chunks),
                    "has_screenshot": True,
                    "screenshot_size": len(screenshot_data) if screenshot_data else 0
                }
            }
            documents_to_insert.append(doc)
            if (i + 1) % 5 == 0 or (i + 1) == len(chunks):
                print(f"   Created embedding {i + 1}/{len(chunks)}")
        if documents_to_insert:
            result = embeddings_collection.insert_many(documents_to_insert)
            print(f"\n   DEBUG: Screenshot stored in MongoDB: document_id={document_id}, chunks={len(documents_to_insert)}\n")
        else:
            print(f"   DEBUG: No documents to insert for {source_url}\n")
    except Exception as e:
        print(f"\nBackground embedding failed for {source_url}: {e}\n")


@app.post("/api/embed-screenshot/")
async def embed_screenshot(
    request: ScreenshotRequest,
    background_tasks: BackgroundTasks,
    authorization: Optional[str] = Header(None),
    x_google_access_token: Optional[str] = Header(None, alias="X-Google-Access-Token"),
):
    """
    Accept screenshot, extract text, then process embeddings in the background.
    Returns immediately so chat and other requests are not blocked.
    For Google Sheets/Docs URLs, pass the user's Google access token in X-Google-Access-Token.
    """
    supabase_token = None
    if authorization and authorization.startswith("Bearer "):
        supabase_token = authorization.split(" ")[1]

    google_token = (x_google_access_token or "").strip() or None
    print(f"\nScreenshot received: {request.source_url} (token present: {bool(supabase_token)}, Google token: {bool(google_token)})\n")

    # Skip embedding for Google OAuth / sign-in pages (they're not useful content and often return 400)
    if "accounts.google.com" in (request.source_url or ""):
        print("   Skipping embed for Google auth page.\n")
        return {
            "status": "skipped",
            "message": "Screenshot accepted but not embedded (Google sign-in page).",
            "source_url": request.source_url,
        }

    text = extract_text_from_url(request.source_url, google_access_token=google_token)
    if not text.strip():
        print(f"No text extracted from URL: {request.source_url}")
        raise HTTPException(
            status_code=400,
            detail="Could not extract text from webpage. The page might require JavaScript or be inaccessible."
        )

    print(f"   Extracted {len(text)} chars; queuing embedding work in background.\n")

    background_tasks.add_task(
        asyncio.to_thread,
        _run_embedding_sync,
        text,
        supabase_token,
        request.source_url,
        request.captured_at,
        request.title,
        request.screenshot_data or "",
    )

    return {
        "status": "accepted",
        "message": "Screenshot accepted; embeddings are being created in the background.",
        "source_url": request.source_url,
    }


GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"


@app.post("/api/google-auth/code")
async def google_auth_code(req: GoogleAuthCodeRequest):
    """
    Exchange Google OAuth authorization code for access_token and refresh_token.
    Requires GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in env.
    """
    client_id = os.getenv("GOOGLE_CLIENT_ID")
    client_secret = os.getenv("GOOGLE_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise HTTPException(status_code=500, detail="Google OAuth not configured")
    data = {
        "code": req.code.strip(),
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": req.redirect_uri.strip(),
        "grant_type": "authorization_code",
    }
    try:
        r = httpx.post(GOOGLE_TOKEN_URL, data=data, timeout=10.0)
        r.raise_for_status()
        body = r.json()
        return {
            "access_token": body.get("access_token"),
            "refresh_token": body.get("refresh_token"),
            "expires_in": body.get("expires_in"),
        }
    except httpx.HTTPStatusError as e:
        raise HTTPException(status_code=e.response.status_code, detail=e.response.text)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/google-auth/refresh")
async def google_auth_refresh(req: GoogleAuthRefreshRequest):
    """
    Exchange refresh_token for a new access_token.
    """
    client_id = os.getenv("GOOGLE_CLIENT_ID")
    client_secret = os.getenv("GOOGLE_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise HTTPException(status_code=500, detail="Google OAuth not configured")
    data = {
        "refresh_token": req.refresh_token.strip(),
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token",
    }
    try:
        r = httpx.post(GOOGLE_TOKEN_URL, data=data, timeout=10.0)
        r.raise_for_status()
        body = r.json()
        return {
            "access_token": body.get("access_token"),
            "expires_in": body.get("expires_in"),
        }
    except httpx.HTTPStatusError as e:
        raise HTTPException(status_code=e.response.status_code, detail=e.response.text)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


def retrieve_context(query: str, k: int = 4, supabase_token: str = None) -> str:
    """
    Retrieve relevant context from vector store and format it as a string
    
    Args:
        query: The search query string
        k: Number of top results to return
        supabase_token: Optional Supabase token to filter documents by user
    
    Returns:
        Formatted string with relevant context from documents
    """
    results = retrieve_documents(query, top_k=k, supabase_token=supabase_token)
    
    if not results:
        return "No relevant context found in the knowledge base."
    
    context_parts = []
    for i, result in enumerate(results, 1):
        context_parts.append(
            f"[Document {i} from {result['filename']}]\n"
            f"{result['text']}\n"
        )
    
    return "\n".join(context_parts)


@app.post("/chat", response_model=ChatResponse)
async def chat(
    request: ChatRequest,
    authorization: Optional[str] = Header(None),
    x_google_access_token: Optional[str] = Header(None, alias="X-Google-Access-Token"),
):
    """
    Chat endpoint that uses Claude agent with RAG (Retrieval Augmented Generation)
    
    Args:
        request: ChatRequest containing:
            - message: The user's message
            - conversation_history: Optional list of previous messages
        authorization: Authorization header containing Bearer token
    
    Returns:
        ChatResponse with the agent's response
    """
    try:
        user_message = request.message
        conversation_history = request.conversation_history or []

        # Extract token from Authorization header
        supabase_token = None
        if authorization and authorization.startswith("Bearer "):
            supabase_token = authorization.split(" ")[1]
        
        google_token = (x_google_access_token or "").strip() or None
        current_tab_url = (request.current_tab_url or "").strip() or None
        print(f"DEBUG: Chat request - Token present: {bool(supabase_token)}, Google token: {bool(google_token)}, Tab URL: {current_tab_url}")

        # Step 1: Retrieve relevant context from vector store (filtered by user token)
        context = retrieve_context(user_message, k=4, supabase_token=supabase_token)
        
        # Step 2: Build the prompt with context
        # System message that instructs the agent to use the retrieved context and return structured output
        system_prompt = """You are a helpful assistant that answers questions based on the provided context from a knowledge base.

When responding, you MUST return your response in the following JSON format.
For "chat_only" or "open_tab": { "action": "chat_only" or "open_tab", "msg": "your response message here" }
For "send_email": { "action": "send_email", "msg": "brief confirmation for the user", "email_to": "recipient@example.com", "email_subject": "Subject line", "email_body": "Draft body text" }

Action rules:
- Use "open_tab" if the user's query requires opening a new browser tab or webpage (e.g., "open YouTube", "search for Python tutorials", "go to github.com")
- When using "open_tab", you MUST include the full URL to open in the "msg" field (e.g. "Opening https://www.wikipedia.org for you." or "Here is the link: https://youtube.com") so the client can open it. Never use "open_tab" with a msg that does not contain a literal https:// or http:// URL.
- Use "send_email" when the user explicitly wants to send, compose, or draft an email (e.g., "send an email to John", "email my team about the meeting", "compose an email to support@example.com"). When using "send_email" you MUST also include "email_to", "email_subject", and "email_body" in the JSON so the client can open a Gmail draft. Do NOT use "send_email" for general questions about email—only when the user wants to actually send one.
- Use "edit_slides" when the user is currently on a Google Slides tab (the current tab URL contains docs.google.com/presentation) AND asks to modify slide layout, alignment, symmetry, positioning, create a new slide, or generate slide content. Put a brief description of what you will do in "msg" (e.g., "Making the text boxes symmetrical on this slide." or "Creating a new slide about AI.").
- Use "chat_only" for all other queries (answering questions, explanations, general conversation)

When answering:
- Use the retrieved context to provide accurate, detailed answers
- If the context contains relevant information, cite it in your response
- If the context doesn't contain enough information to answer the question, say so
- You can also use your general knowledge, but prioritize the provided context
- Be concise but thorough
- Always return valid JSON with both "action" and "msg" fields"""
        
        messages = [SystemMessage(content=system_prompt)]
        
        # Add conversation history
        for msg in conversation_history:
            if msg.role == "user":
                messages.append(HumanMessage(content=msg.content))
            elif msg.role == "assistant":
                messages.append(AIMessage(content=msg.content))
        
        # Add current user message with context
        tab_context = f"\n            The user's current browser tab URL is: {current_tab_url}" if current_tab_url else ""
        user_prompt = f"""Context from knowledge base:
            {context}

            User question: {user_message}{tab_context}

            The current time of the user request is: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

            Please answer the user's question using the context above when available."""
        
        messages.append(HumanMessage(content=user_prompt))
        
        # Step 3: Get response from LLM (the "agent")
        response = llm.invoke(messages)
        response_text = response.content if hasattr(response, 'content') else str(response)
        
        # Step 4: Parse structured response from LLM
        # Try to extract JSON from the response
        action = ActionType.CHAT_ONLY  # Default action
        msg = response_text  # Default message
        email_url = None

        try:
            # Try to find JSON in the response (handles cases where LLM adds extra text)
            json_match = re.search(r'\{[^{}]*"action"[^{}]*"msg"[^{}]*\}', response_text, re.DOTALL)
            if json_match:
                parsed = json.loads(json_match.group())
                action_str = parsed.get("action", "chat_only").lower()
                msg = parsed.get("msg", response_text)

                # Validate and set action
                if action_str == "open_tab":
                    action = ActionType.OPEN_TAB
                elif action_str == "send_email":
                    action = ActionType.SEND_EMAIL
                elif action_str == "edit_slides":
                    action = ActionType.EDIT_SLIDES
                else:
                    action = ActionType.CHAT_ONLY
            else:
                # If no JSON found, try parsing the entire response as JSON
                parsed = json.loads(response_text)
                action_str = parsed.get("action", "chat_only").lower()
                msg = parsed.get("msg", response_text)

                if action_str == "open_tab":
                    action = ActionType.OPEN_TAB
                elif action_str == "send_email":
                    action = ActionType.SEND_EMAIL
                elif action_str == "edit_slides":
                    action = ActionType.EDIT_SLIDES
                else:
                    action = ActionType.CHAT_ONLY

            # Build Gmail compose URL when action is send_email
            if action == ActionType.SEND_EMAIL:
                email_to = (parsed.get("email_to") or "").strip()
                email_subject = (parsed.get("email_subject") or "").strip()
                email_body = (parsed.get("email_body") or "").strip()
                # Truncate body to avoid URL length limits (~2000 chars safe)
                max_body_len = 1500
                if len(email_body) > max_body_len:
                    email_body = email_body[:max_body_len] + "..."
                if email_to or email_subject or email_body:
                    email_url = "https://mail.google.com/mail/?view=cm&fs=1&" + urlencode({
                        "to": email_to,
                        "su": email_subject,
                        "body": email_body,
                    })
        except (json.JSONDecodeError, KeyError, AttributeError) as e:
            # If JSON parsing fails, use the raw response as message with default action
            print(f"DEBUG: Failed to parse structured response, using default: {str(e)}")
            print(f"   Raw response: {response_text[:200]}...")
            msg = response_text
            action = ActionType.CHAT_ONLY

        # When action is edit_slides, execute the slides module
        if action == ActionType.EDIT_SLIDES:
            slides_result = handle_edit_slides(current_tab_url, user_message, google_token)
            msg = slides_result

        # Debug: log action and, for open_tab/send_email, the message content
        print(f"   Action: {action.value}")
        if action == ActionType.OPEN_TAB:
            print(f"   OPEN_TAB: msg (first 300 chars): {msg[:300] if msg else '(empty)'}")
        if action == ActionType.SEND_EMAIL:
            print(f"   SEND_EMAIL: msg (first 300 chars): {msg[:300] if msg else '(empty)'}")
            if email_url:
                print(f"   SEND_EMAIL: email_url generated")
        if action == ActionType.EDIT_SLIDES:
            print(f"   EDIT_SLIDES: msg (first 300 chars): {msg[:300] if msg else '(empty)'}")

        return ChatResponse(action=action, msg=msg, email_url=email_url)
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)