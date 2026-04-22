from dataclasses import dataclass
import os

import streamlit as st

try:
    from openai import AzureOpenAI, OpenAI
except Exception:
    AzureOpenAI = None
    OpenAI = None


DEFAULT_API_VERSION = "2024-02-01"
DEFAULT_AZURE_OPENAI_ENDPOINT = "https://mwlab-azoai.openai.azure.com/"


@dataclass
class AzureOpenAIConfig:
    endpoint: str = ""
    api_key: str = ""
    deployment: str = ""
    api_version: str = DEFAULT_API_VERSION

    @property
    def ready(self) -> bool:
        return bool(self.endpoint and self.api_key and self.deployment)

    @property
    def endpoint_base(self) -> str:
        return self.endpoint.rstrip("/")


def _secret_or_env(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return str(value or os.getenv(name, default) or "").strip()


def get_azure_openai_config(prefix: str = "budg") -> AzureOpenAIConfig:
    endpoint_key = f"{prefix}_azure_openai_endpoint"
    api_key_key = f"{prefix}_azure_openai_api_key"
    deployment_key = f"{prefix}_azure_openai_deployment"
    api_version_key = f"{prefix}_azure_openai_api_version"

    endpoint = st.session_state.get(endpoint_key) or _secret_or_env(
        "AZURE_OPENAI_ENDPOINT",
        DEFAULT_AZURE_OPENAI_ENDPOINT,
    )
    api_key = st.session_state.get(api_key_key) or _secret_or_env("AZURE_OPENAI_API_KEY")
    deployment = st.session_state.get(deployment_key) or _secret_or_env("AZURE_OPENAI_DEPLOYMENT_NAME")
    api_version = st.session_state.get(api_version_key) or _secret_or_env(
        "AZURE_OPENAI_API_VERSION",
        DEFAULT_API_VERSION,
    )
    return AzureOpenAIConfig(
        endpoint=str(endpoint).strip(),
        api_key=str(api_key).strip(),
        deployment=str(deployment).strip(),
        api_version=str(api_version).strip() or DEFAULT_API_VERSION,
    )


def render_azure_openai_settings(prefix: str = "budg") -> AzureOpenAIConfig:
    config = get_azure_openai_config(prefix)

    st.markdown(
        """
        <div style="padding:0.9rem 1rem;border:1px solid rgba(49,51,63,0.18);border-radius:16px;
                    background:linear-gradient(135deg, rgba(16,124,94,0.10), rgba(0,92,151,0.08));">
          <div style="font-size:0.82rem;letter-spacing:0.08em;text-transform:uppercase;
                      color:var(--text-color-secondary, #6b7280);margin-bottom:0.25rem;">
            AI Copilot
          </div>
          <div style="font-size:1.05rem;font-weight:600;margin-bottom:0.2rem;">
            Azure OpenAI for management commentary and portfolio Q&amp;A
          </div>
          <div style="font-size:0.92rem;color:var(--text-color-secondary, #6b7280);">
            Uses your own Azure deployment. Session values override secrets and environment variables.
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("Azure OpenAI settings", expanded=not config.ready):
        col1, col2 = st.columns(2)
        with col1:
            st.text_input(
                "Endpoint",
                value=config.endpoint,
                key=f"{prefix}_azure_openai_endpoint",
                placeholder=DEFAULT_AZURE_OPENAI_ENDPOINT,
            )
            st.text_input(
                "Deployment name",
                value=config.deployment,
                key=f"{prefix}_azure_openai_deployment",
                placeholder="gpt-4o-mini-prod",
            )
        with col2:
            st.text_input(
                "API key",
                value=config.api_key,
                key=f"{prefix}_azure_openai_api_key",
                type="password",
                placeholder="Paste Azure OpenAI key",
            )
            st.text_input(
                "API version",
                value=config.api_version,
                key=f"{prefix}_azure_openai_api_version",
                placeholder=DEFAULT_API_VERSION,
            )

        config = get_azure_openai_config(prefix)
        if config.ready:
            st.success(f"Azure OpenAI is ready. Deployment: `{config.deployment}`")
        else:
            st.info(
                "Add `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_API_KEY`, "
                "`AZURE_OPENAI_DEPLOYMENT_NAME`, and optionally `AZURE_OPENAI_API_VERSION` "
                "in `.streamlit/secrets.toml` or enter them here for this session."
            )

        if AzureOpenAI is None:
            st.warning("The `openai` package is not installed yet. Add it to the environment to enable AI.")
        elif config.ready:
            if st.button("Validate Azure AI", key=f"{prefix}_validate_azure_openai", use_container_width=True):
                try:
                    reply = run_azure_openai_text(
                        config,
                        [{"role": "user", "content": "Reply with exactly: Azure AI connected"}],
                        max_tokens=20,
                        temperature=0,
                    )
                    st.success(f"Connection successful: {reply}")
                except Exception as e:
                    st.error(f"Validation failed: {e}")

    return config


def build_azure_openai_client(config: AzureOpenAIConfig):
    if AzureOpenAI is None:
        raise RuntimeError("The `openai` package is not installed.")
    if not config.ready:
        raise RuntimeError("Azure OpenAI is not configured.")
    return AzureOpenAI(
        azure_endpoint=config.endpoint_base,
        api_key=config.api_key,
        api_version=config.api_version or DEFAULT_API_VERSION,
    )


def build_azure_openai_v1_client(config: AzureOpenAIConfig):
    if OpenAI is None:
        raise RuntimeError("The `openai` package is not installed.")
    if not config.ready:
        raise RuntimeError("Azure OpenAI is not configured.")
    return OpenAI(
        api_key=config.api_key,
        base_url=f"{config.endpoint_base}/openai/v1/",
    )


def _messages_to_text(messages: list[dict]) -> tuple[str | None, str]:
    system_parts = []
    user_parts = []
    for message in messages:
        role = str(message.get("role", "")).strip().lower()
        content = str(message.get("content", "")).strip()
        if not content:
            continue
        if role == "system":
            system_parts.append(content)
        else:
            user_parts.append(f"{role or 'user'}: {content}")
    instructions = "\n\n".join(system_parts).strip() or None
    prompt = "\n\n".join(user_parts).strip()
    return instructions, prompt


def run_azure_openai_text(
    config: AzureOpenAIConfig,
    messages: list[dict],
    max_tokens: int = 800,
    temperature: float = 0.2,
) -> str:
    errors = []

    try:
        chat_client = build_azure_openai_client(config)
        response = chat_client.chat.completions.create(
            model=config.deployment,
            messages=messages,
            max_tokens=max_tokens,
            temperature=temperature,
        )
        content = response.choices[0].message.content
        return content.strip() if isinstance(content, str) else str(content).strip()
    except Exception as exc:
        errors.append(f"chat.completions: {exc}")

    try:
        v1_client = build_azure_openai_v1_client(config)
        instructions, prompt = _messages_to_text(messages)
        response_kwargs = {
            "model": config.deployment,
            "instructions": instructions,
            "input": prompt,
            "max_output_tokens": max_tokens,
        }
        if temperature is not None:
            response_kwargs["temperature"] = temperature
        response = v1_client.responses.create(**response_kwargs)
        text = getattr(response, "output_text", "")
        if text:
            return text.strip()
        return str(response).strip()
    except Exception as exc:
        first_error = str(exc)
        if "temperature" in first_error.lower() and "not supported" in first_error.lower():
            try:
                response = v1_client.responses.create(
                    model=config.deployment,
                    instructions=instructions,
                    input=prompt,
                    max_output_tokens=max_tokens,
                )
                text = getattr(response, "output_text", "")
                if text:
                    return text.strip()
                return str(response).strip()
            except Exception as retry_exc:
                errors.append(f"responses: {retry_exc}")
        else:
            errors.append(f"responses: {exc}")

    raise RuntimeError(
        " ; ".join(errors)
        + " ; hint: for Azure chat deployments, try API version 2024-02-01. "
        + "The Responses API is not enabled in West Europe."
    )
