"""
Shared LLM factory for all agents.

Provider selection (in order of priority):
  1. AGENT_PROVIDER env var (explicit: "grok", "openai", "anthropic")
  2. Auto-detect from which API key is present
     ANTHROPIC_API_KEY present → Anthropic (Claude Opus — default)
     XAI_API_KEY present       → Grok
     OPENAI_API_KEY present    → OpenAI
"""
import os
import litellm
from crewai import LLM

# Drop unsupported parameters (e.g. reasoning models like grok-4-1-fast-reasoning
# reject 'stop' sequences that CrewAI sends by default)
litellm.drop_params = True


# Default models per provider
_DEFAULTS = {
    "grok":      "grok-4-1-fast-non-reasoning",
    "openai":    "gpt-4o-mini",
    "anthropic": "claude-opus-4-6",
}


def get_llm() -> LLM:
    """Return an LLM instance based on AGENT_PROVIDER / available API keys."""

    provider  = os.getenv("AGENT_PROVIDER", "").lower().strip()
    xai_key   = os.getenv("XAI_API_KEY")   or os.getenv("GROK_API_KEY")
    openai_key = os.getenv("OPENAI_API_KEY")
    anthropic_key = os.getenv("ANTHROPIC_API_KEY")

    # Auto-detect if provider not set — Anthropic preferred
    if not provider:
        if anthropic_key:
            provider = "anthropic"
        elif xai_key:
            provider = "grok"
        elif openai_key:
            provider = "openai"
        else:
            provider = "anthropic"

    model = os.getenv("AGENT_MODEL", _DEFAULTS.get(provider, "claude-opus-4-6"))

    if provider == "grok":
        if not xai_key:
            raise EnvironmentError(
                "AGENT_PROVIDER=grok but XAI_API_KEY / GROK_API_KEY is not set in .env"
            )
        return LLM(
            model=f"openai/{model}",
            base_url="https://api.x.ai/v1",
            api_key=xai_key,
        )

    if provider == "openai":
        if not openai_key:
            raise EnvironmentError(
                "AGENT_PROVIDER=openai but OPENAI_API_KEY is not set in .env"
            )
        return LLM(
            model=f"openai/{model}",
            api_key=openai_key,
        )

    # Anthropic (default)
    if not anthropic_key:
        raise EnvironmentError(
            "No API key found. Set XAI_API_KEY, OPENAI_API_KEY, or ANTHROPIC_API_KEY in .env"
        )
    return LLM(
        model=f"anthropic/{model}",
        api_key=anthropic_key,
        max_tokens=4096,
    )


def get_litellm_config() -> dict:
    """
    Return raw litellm call parameters for direct (non-CrewAI) LLM calls.
    Used by the chat mode conversation loop.

    Returns dict with keys: model, api_key, and optionally base_url.
    Pass as **get_litellm_config() to litellm.completion().
    """
    provider  = os.getenv("AGENT_PROVIDER", "").lower().strip()
    xai_key   = os.getenv("XAI_API_KEY")   or os.getenv("GROK_API_KEY")
    openai_key = os.getenv("OPENAI_API_KEY")
    anthropic_key = os.getenv("ANTHROPIC_API_KEY")

    if not provider:
        if anthropic_key:
            provider = "anthropic"
        elif xai_key:
            provider = "grok"
        elif openai_key:
            provider = "openai"
        else:
            provider = "anthropic"

    model = os.getenv("AGENT_MODEL", _DEFAULTS.get(provider, "claude-opus-4-6"))

    if provider == "grok":
        return {
            "model": f"openai/{model}",
            "api_key": xai_key,
            "base_url": "https://api.x.ai/v1",
        }
    if provider == "openai":
        return {
            "model": f"openai/{model}",
            "api_key": openai_key,
        }
    # Anthropic
    return {
        "model": f"anthropic/{model}",
        "api_key": anthropic_key,
        "max_tokens": 4096,
    }
