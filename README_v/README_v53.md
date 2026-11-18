# Desktop AI v5.3 — Multi-Provider Edition

Apply over v5.2. Replace main.py and DesktopAI.iss. No new dependencies.

In API Settings enter an HTTPS OpenAI-compatible base URL, API key, model ID, RPM and timeout. Click Discover models to fetch the first model-list page; select a model or type its ID. Test native tool calling before use. Vision support is not tested by this check. Save profile stores a named configuration; Load activates it in the form. Save Settings persists the active configuration for startup. Delete removes only the saved profile, not active settings or the provider's remote key. Remember API key controls whether the profile stores DPAPI-encrypted credentials. Unchecked profiles require re-entry. Existing startup settings remain compatible.

Profiles live in %LOCALAPPDATA%/DesktopAI/providers.json. Never publish this file. Shared task memory can be sent to whichever provider is active; reset memory before switching if needed. Model discovery and connection tests are not included in the agent's request counter. Agent rate limiting remains shared across providers. HTTPS OpenAI-compatible APIs only; native Anthropic/Gemini APIs are not supported. Model discovery may be unavailable on some gateways. Rebuild using build_windows.cmd and optionally compile DesktopAI.iss.

Syntax checked; Windows GUI, DPAPI, live endpoints and EXE build not tested here.
