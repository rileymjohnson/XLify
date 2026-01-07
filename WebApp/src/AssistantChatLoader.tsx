import React, { Suspense } from 'react'

type LoaderProps = {}

const ChatApp = React.lazy(async () => {
  try {
    // Dynamically import library (styles can be added separately in index.html or a global css import)
    const mod: any = await import('@assistant-ui/react')

    const { AssistantRuntimeProvider, useLocalRuntime, ThreadPrimitive, ComposerPrimitive } = mod as any

    const Comp: React.FC<LoaderProps> = ({}) => {
      // Minimal local runtime without a chat model adapter (UI only)
      const runtime = useLocalRuntime ? useLocalRuntime(undefined, { initialMessages: [] }) : null
      const [msgs, setMsgs] = React.useState<{ role: 'user' | 'assistant' | 'echo' | 'error'; text: string }[]>([])
      const [isSending, setIsSending] = React.useState(false)
      const [status, setStatus] = React.useState<string>('')
      const chatRef = React.useRef<HTMLDivElement | null>(null)
      React.useEffect(() => {
        try {
          const el = chatRef.current
          if (!el) return
          requestAnimationFrame(() => { try { el.scrollTop = el.scrollHeight } catch {} })
        } catch {}
      }, [msgs, status])
      const inputRef = React.useRef<HTMLTextAreaElement | null>(null)
      const autoResize = () => {
        const ta = inputRef.current
        if (!ta) return
        ta.style.height = 'auto'
        const max = 160
        const next = Math.min(ta.scrollHeight, max)
        ta.style.height = next + 'px'
        ta.style.overflowY = ta.scrollHeight > max ? 'auto' : 'hidden'
      }
      const handleKeyDown: React.KeyboardEventHandler<HTMLTextAreaElement> = (e) => {
        if (e.key === 'Enter' && !e.shiftKey && !e.altKey && !e.ctrlKey && !e.metaKey) {
          e.preventDefault()
          const form = (e.currentTarget as HTMLTextAreaElement).form as HTMLFormElement | null
          try { form?.requestSubmit() } catch { /* fallback */ form?.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true })) }
        }
      }
      React.useEffect(() => {
        const onMsg = (ev: MessageEvent) => {
          try {
            const data = typeof ev.data === 'string' ? JSON.parse(ev.data) : ev.data
            if (data?.type) {
              if (data.type === 'status') {
                // Update live status banner; do not add to conversation
                const text = typeof data.text === 'string' ? data.text : ''
                setStatus(text || '')
                return
              }
              const text = data.text ?? (data.data ? JSON.stringify(data.data) : undefined)
              if (typeof text === 'string') setMsgs(m => [...m, { role: (data.type as any) ?? 'echo', text }])
              if (data.type === 'assistant' || data.type === 'error') {
                setIsSending(false)
                // Clear status when assistant/error completes
                setStatus('')
              }
            }
          } catch { /* ignore */ }
        }
        ;(window as any).chrome?.webview?.addEventListener('message', onMsg)
        return () => (window as any).chrome?.webview?.removeEventListener('message', onMsg)
      }, [])
      if (!AssistantRuntimeProvider || !ThreadPrimitive || !ComposerPrimitive || !runtime) {
        const keys = Object.keys(mod || {})
        return (
          <div className="p-3 text-sm text-gray-700">
            <div>assistant-ui loaded. Exports: {keys.join(', ') || 'none'}</div>
            <div>Install and configure a backend adapter to enable chat.</div>
          </div>
        )
      }
      return (
        <AssistantRuntimeProvider runtime={runtime}>
          <div className="flex h-full w-full flex-col">
            <div ref={chatRef} className="min-h-0 flex-1 overflow-auto p-3 space-y-2">
              {msgs.map((m, i) => {
                const isUser = m.role === 'user'
                const isError = m.role === 'error'
                const bubble = isUser
                  ? 'bg-blue-600 text-white'
                  : isError
                  ? 'bg-red-50 text-red-700 border border-red-200'
                  : 'bg-gray-100 text-gray-900'
                const align = isUser ? 'justify-end' : 'justify-start'
                const radius = isUser ? 'rounded-2xl rounded-br-sm' : 'rounded-2xl rounded-bl-sm'
                return (
                  <div key={i} className={`w-full flex ${align}`}>
                    <div className={`max-w-[80%] px-3 py-2 ${bubble} ${radius} shadow-sm whitespace-pre-wrap break-words`}>{m.text}</div>
                  </div>
                )
              })}
              {Boolean(status) && (
                <div className="w-full flex justify-start">
                  <div className="max-w-[80%] px-3 py-2 bg-gray-100 text-gray-600 rounded-2xl rounded-bl-sm shadow-sm flex items-center gap-2">
                    <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z"></path>
                    </svg>
                    <span>{status}</span>
                  </div>
                </div>
              )}
            </div>
            {/* Temporary custom composer to avoid runtime errors until a chat model adapter is configured */}
            <div className="border-t border-gray-200 p-2">
              <form className="flex items-center gap-2" onSubmit={(e) => {
                e.preventDefault();
                const form = e.currentTarget as HTMLFormElement;
                const input = form.querySelector('textarea[name="msg"]') as HTMLTextAreaElement | null;
                const text = input?.value?.trim();
                if (isSending) {
                  return;
                }
                if (text) {
                  setMsgs(m => [...m, { role: 'user', text }]);
                  setIsSending(true);
                  setStatus('Thinking…');
                  (window as any).chrome?.webview?.postMessage({ type: 'user', text });
                  if (input) { input.value = ''; autoResize(); }
                }
              }}>
                <textarea
                  name="msg"
                  ref={inputRef}
                  placeholder="Ask anything..."
                  disabled={isSending}
                  rows={1}
                  onInput={autoResize}
                  onKeyDown={handleKeyDown}
                  className="flex-1 rounded-md border border-gray-300 px-3 py-2 outline-none focus:border-blue-500 disabled:opacity-50 disabled:cursor-not-allowed resize-none leading-5"
                  style={{ maxHeight: 160, overflowY: 'hidden' }}
                />
                <button type="submit" aria-label="Send" disabled={isSending} className="inline-flex items-center gap-2 rounded-md bg-blue-600 px-3 py-2 text-white hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed">
                  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-4 w-4">
                    <path d="M2.01 21 23 12 2.01 3 2 10l15 2-15 2z" />
                  </svg>
                  <span className="sr-only">Send</span>
                </button>
              </form>
              {/* helper text removed */}
            </div>
          </div>
        </AssistantRuntimeProvider>
      )
    }

    return { default: Comp }
  } catch (e: any) {
    const Err = () => <div style={{ padding: 12 }}>Failed to load @assistant-ui/react: {String(e?.message || e)}</div>
    return { default: Err }
  }
})

export default function AssistantChatLoader(props: LoaderProps) {
  return (
    <Suspense fallback={<div style={{ padding: 12 }}>Loading assistant UI…</div>}>
      <ChatApp {...props} />
    </Suspense>
  )
}
