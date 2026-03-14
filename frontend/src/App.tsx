import { useState, useRef, useEffect } from 'react'
import { useAuth } from './contexts/AuthContext'
import { authenticatedFetch } from './utils/api'
import Login from './components/Login'
import SignUp from './components/SignUp'
import { Message } from './types'
import { openTabFromMessage, openEmailCompose } from './Tools'
import { getGoogleTokens, clearGoogleTokens, getValidGoogleAccessToken } from './utils/googleAuth'
import './App.css'

const API_URL = 'http://localhost:8000'
const GOOGLE_SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/documents.readonly https://www.googleapis.com/auth/presentations'

function App() {
  const { user, loading, signOut } = useAuth()
  const [showSignUp, setShowSignUp] = useState(false)

  // Show loading state while checking auth
  if (loading) {
    return (
      <div className="app">
        <div className="container">
          <div className="loading-container">
            <div className="spinner"></div>
            <p>Loading...</p>
          </div>
        </div>
      </div>
    )
  }

  // Show auth pages if not logged in
  if (!user) {
    return showSignUp ? (
      <SignUp onSwitchToLogin={() => setShowSignUp(false)} />
    ) : (
      <Login onSwitchToSignUp={() => setShowSignUp(true)} />
    )
  }

  // Show chatbot if logged in
  return <Chatbot signOut={signOut} />
}

interface ChatbotProps {
  signOut: () => Promise<{ error: any }>
}

function Chatbot({ signOut }: ChatbotProps) {
  const [messages, setMessages] = useState<Message[]>([
    { role: 'assistant', content: 'Hello! How can I help you today?' }
  ])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [googleConnected, setGoogleConnected] = useState(false)
  const [googleConnecting, setGoogleConnecting] = useState(false)
  const [googleError, setGoogleError] = useState<string | null>(null)
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const refreshGoogleState = () => {
    getGoogleTokens().then((t) => setGoogleConnected(!!t.access_token))
  }

  useEffect(() => {
    refreshGoogleState()
    const c = typeof chrome !== 'undefined' ? chrome : (window as any).chrome
    if (c?.storage?.onChanged?.addListener) {
      const listener = () => refreshGoogleState()
      c.storage.onChanged.addListener(listener)
      return () => c.storage.onChanged.removeListener(listener)
    }
  }, [])

  useEffect(() => {
    const onFocus = () => refreshGoogleState()
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [])

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }

  useEffect(() => {
    scrollToBottom()
  }, [messages])

  useEffect(() => {
    inputRef.current?.focus()
  }, [])

  const sendMessage = async () => {
    const message = input.trim()
    if (!message || loading) return

    setInput('')
    setLoading(true)

    // Add user message
    const newUserMessage: Message = { role: 'user', content: message }
    setMessages(prev => [...prev, newUserMessage])

    try {
      // Prepare conversation history (excluding the welcome message)
      const conversationHistory: Message[] = messages
        .filter(msg => msg.role !== 'assistant' || msg.content !== 'Hello! How can I help you today?')
        .map(msg => ({
          role: msg.role,
          content: msg.content
        }))

      // Get active tab URL for context (e.g. Google Slides editing)
      let currentTabUrl: string | undefined
      const c = typeof chrome !== 'undefined' ? chrome : (window as any).chrome
      if (c?.tabs?.query) {
        try {
          const [tab] = await c.tabs.query({ active: true, currentWindow: true })
          currentTabUrl = tab?.url
        } catch (_) {}
      }

      // Get Google token for Slides API access
      const googleToken = await getValidGoogleAccessToken(API_URL)

      const headers: Record<string, string> = {}
      if (googleToken) {
        headers['X-Google-Access-Token'] = googleToken
      }

      const response = await authenticatedFetch(`${API_URL}/chat`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          message: message,
          conversation_history: conversationHistory,
          current_tab_url: currentTabUrl ?? null,
        })
      })

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`)
      }

      const data = await response.json()
      // ChatResponse: { action: 'chat_only' | 'open_tab' | 'send_email', msg: string, email_url?: string }
      const content = data.msg ?? data.response ?? ''
      console.log('[Chat] response action:', data.action, 'msg length:', content?.length)
      setMessages(prev => [...prev, { role: 'assistant', content }])
      if (data.action === 'open_tab') {
        console.log('[Chat] OPEN_TAB: calling openTabFromMessage with content:', content?.slice(0, 200))
        openTabFromMessage(content)
      }
      if (data.action === 'send_email' && data.email_url) {
        console.log('[Chat] SEND_EMAIL: opening Gmail compose')
        openEmailCompose(data.email_url)
      }
      if (data.action === 'edit_slides') {
        console.log('[Chat] EDIT_SLIDES: slide updated')
      }
    } catch (error: any) {
      console.error('Error:', error)
      setMessages(prev => [...prev, { 
        role: 'assistant', 
        content: `Error: ${error.message}. Make sure the backend is running on ${API_URL}` 
      }])
    } finally {
      setLoading(false)
      inputRef.current?.focus()
    }
  }

  const handleKeyPress = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      sendMessage()
    }
  }

  const handleSignOut = async () => {
    await signOut()
  }

  const handleConnectGoogle = async () => {
    setGoogleError(null)
    setGoogleConnecting(true)
    const clientId = (import.meta as any).env?.VITE_GOOGLE_CLIENT_ID ?? ''
    const c = typeof chrome !== 'undefined' ? chrome : (window as any).chrome
    if (!c?.identity?.launchWebAuthFlow || !c?.runtime?.id) {
      setGoogleConnecting(false)
      setGoogleError('Go to chrome://extensions → click the refresh icon on this extension, then try Connect Google again.')
      return
    }
    if (!clientId) {
      setGoogleConnecting(false)
      setGoogleError('Add VITE_GOOGLE_CLIENT_ID to frontend/.env and run: npm run build, then reload the extension.')
      return
    }
    try {
      const redirectUri = `https://${c.runtime.id}.chromiumapp.org/`
      const authUrl =
        'https://accounts.google.com/o/oauth2/v2/auth?' +
        new URLSearchParams({
          client_id: clientId,
          redirect_uri: redirectUri,
          response_type: 'code',
          scope: GOOGLE_SCOPES,
          access_type: 'offline',
          prompt: 'consent',
        }).toString()
      const redirectUrl = await c.identity.launchWebAuthFlow({
        url: authUrl,
        interactive: true,
      })
      const parsed = new URL(redirectUrl)
      const code = parsed.searchParams.get('code')
      if (!code) throw new Error('No code in redirect')
      // Send code to background so it does the exchange and writes to storage (persists even if popup closes)
      const response = await new Promise<{ success: boolean; error?: string }>((resolve) => {
        c.runtime.sendMessage(
          { type: 'GOOGLE_AUTH_SAVE', code, redirect_uri: redirectUri },
          (r: { success: boolean; error?: string } | undefined) => resolve(r || { success: false })
        )
      })
      if (!response.success) throw new Error(response.error || 'Failed to save tokens')
      const check = await getGoogleTokens()
      setGoogleConnected(!!check.access_token)
      if (!check.access_token) setGoogleError('Saved in background; reopen the popup to see Disconnect Google.')
    } catch (e: any) {
      const msg = e?.message || String(e)
      if (msg.includes('Authorization page could not be loaded') || msg.includes('redirect_uri')) {
        setGoogleError('Add redirect URI in Google Cloud: https://YOUR_EXTENSION_ID.chromiumapp.org/')
      } else if (msg.includes('canceled') || msg.includes('cancelled') || msg.includes('user closed')) {
        setGoogleError('Sign-in was cancelled.')
      } else {
        setGoogleError(msg.length > 80 ? msg.slice(0, 80) + '…' : msg)
      }
      console.error('Connect Google failed:', e)
    } finally {
      setGoogleConnecting(false)
    }
  }

  const handleDisconnectGoogle = async () => {
    await clearGoogleTokens()
    setGoogleConnected(false)
    setGoogleError(null)
  }

  return (
    <div className="app">
      <div className="container">
        <div className="header">
          <h2>🤖 Dex2</h2>
          <div className="header-actions">
            {googleConnected ? (
              <button type="button" onClick={handleDisconnectGoogle} className="logout-button">
                Disconnect Google
              </button>
            ) : (
              <button
                type="button"
                onClick={handleConnectGoogle}
                disabled={googleConnecting}
                className="logout-button"
              >
                {googleConnecting ? 'Connecting…' : 'Connect Google'}
              </button>
            )}
            <button onClick={handleSignOut} className="logout-button">
              Sign out
            </button>
          </div>
        </div>
        {googleError && (
          <div className="google-error" role="alert">
            {googleError}
          </div>
        )}
        
        <div className="messages" id="messages">
          {messages.map((message, index) => (
            <div key={index} className={`message ${message.role}`}>
              <div className="message-content">
                {message.content}
              </div>
            </div>
          ))}
          
          {loading && (
            <div className="message assistant">
              <div className="message-content loading">
                Thinking<span className="dots">...</span>
              </div>
            </div>
          )}
          
          <div ref={messagesEndRef} />
        </div>

        <div className="input-container">
          <input
            ref={inputRef}
            type="text"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyPress={handleKeyPress}
            placeholder="Type your message..."
            disabled={loading}
            className="input-field"
          />
          <button
            onClick={sendMessage}
            disabled={loading || !input.trim()}
            className="send-button"
          >
            {loading ? '⏳' : '📤'}
          </button>
        </div>
      </div>
    </div>
  )
}

export default App
