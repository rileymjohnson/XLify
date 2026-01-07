import React from 'react'
import AssistantChatLoader from './AssistantChatLoader'
import ApiKeyGate from './ApiKeyGate'

export default function App() {
  const [hasKey, setHasKey] = React.useState<boolean | null>(null)
  const [showChange, setShowChange] = React.useState(false)

  React.useEffect(() => {
    const onMsg = (ev: MessageEvent) => {
      try {
        const data = typeof ev.data === 'string' ? JSON.parse(ev.data) : ev.data
        if (data?.type === 'hasApiKey') {
          setHasKey(String(data.text || '').toLowerCase() === 'true')
        }
        if (data?.type === 'saveApiKey' && (data.text === 'ok')) {
          setHasKey(true)
          setShowChange(false)
        }
        if (data?.type === 'clearApiKey' && (data.text === 'ok')) {
          setHasKey(false)
          setShowChange(false)
        }
      } catch {}
    }
    ;(window as any).chrome?.webview?.addEventListener('message', onMsg)
    ;(window as any).chrome?.webview?.postMessage({ type: 'hasApiKey' })
    return () => (window as any).chrome?.webview?.removeEventListener('message', onMsg)
  }, [])

  if (hasKey === null) {
    return <div className="h-full w-full flex items-center justify-center text-gray-500">Loading…</div>
  }

  if (!hasKey) {
    return (
      <div style={{ height: '100%', width: '100%' }}>
        <ApiKeyGate onSubmit={(apiKey) => {
          try { (window as any).chrome?.webview?.postMessage({ type: 'saveApiKey', apiKey }) } catch {}
        }} />
      </div>
    )
  }

  return (
    <div className="h-full w-full flex flex-col">
      <div className="flex items-center gap-2 border-b border-gray-200 px-3 py-2 text-sm bg-white">
        <div className="font-semibold">XLify</div>
        <div className="ml-auto flex items-center gap-2">
          <button className="rounded-md border border-gray-300 px-2 py-1 hover:bg-gray-50" onClick={() => setShowChange(true)}>Change API Key</button>
          <button className="rounded-md border border-gray-300 px-2 py-1 text-red-700 hover:bg-red-50" onClick={() => {
            try { (window as any).chrome?.webview?.postMessage({ type: 'clearApiKey' }) } catch {}
          }}>Remove Key</button>
        </div>
      </div>
      <div className="flex-1 min-h-0">
        <AssistantChatLoader />
      </div>
      {showChange && (
        <div className="absolute inset-0 z-10 flex items-center justify-center bg-black/20">
          <ApiKeyGate
            title="Change your OpenAI API Key"
            onSubmit={(apiKey) => {
              try { (window as any).chrome?.webview?.postMessage({ type: 'saveApiKey', apiKey }) } catch {}
            }}
            onCancel={() => setShowChange(false)}
          />
        </div>
      )}
    </div>
  )
}
