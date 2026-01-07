import React from 'react'

type Props = {
  onSubmit: (apiKey: string) => void
  onCancel?: () => void
  title?: string
}

export default function ApiKeyGate({ onSubmit, onCancel, title }: Props) {
  const [key, setKey] = React.useState('')
  const [error, setError] = React.useState<string | null>(null)

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    const trimmed = key.trim()
    if (!trimmed) {
      setError('API key is required')
      return
    }
    // Optional: basic shape check (OpenAI keys often start with sk-)
    if (!/^sk-/.test(trimmed)) {
      // Still allow submission, but warn
      setError('Warning: Key does not look like a standard OpenAI key (sk-...)')
      // Continue; remove the early return if you want to strictly enforce
    } else {
      setError(null)
    }
    onSubmit(trimmed)
  }

  return (
    <div className="h-full w-full flex items-center justify-center bg-white">
      <div className="w-full max-w-md rounded-lg border border-gray-200 shadow-sm p-6 bg-white">
        <h1 className="text-lg font-semibold mb-2">{title ?? 'Enter your OpenAI API Key'}</h1>
        <p className="text-sm text-gray-600 mb-4">
          Bring your own API key to use XLify. Your key will be used to call the OpenAI API from the host application.
        </p>
        <form onSubmit={handleSubmit} className="space-y-3">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1" htmlFor="apiKey">API Key</label>
            <input
              id="apiKey"
              type="password"
              className="w-full rounded-md border border-gray-300 px-3 py-2 outline-none focus:border-blue-500"
              placeholder="sk-..."
              value={key}
              onChange={(e) => setKey(e.target.value)}
              autoFocus
            />
          </div>
          {error && <div className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-2 py-1">{error}</div>}
          <div className="pt-1 flex gap-2">
            <button type="submit" className="rounded-md bg-blue-600 px-3 py-2 text-white hover:bg-blue-700">Use Key</button>
            {onCancel && (
              <button type="button" onClick={onCancel} className="rounded-md border border-gray-300 px-3 py-2 text-gray-700 hover:bg-gray-50">Cancel</button>
            )}
          </div>
        </form>
        <div className="mt-4 text-xs text-gray-500">
          You can create a key in your OpenAI account dashboard.
        </div>
      </div>
    </div>
  )
}
