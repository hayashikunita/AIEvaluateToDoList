from __future__ import annotations

import json
import os
import time
from typing import Dict, Any, List

from openai import OpenAI
from dotenv import load_dotenv


JSON_KEYS = [
    "AI_工数",
    "AI_Level",
    "AI_NextAction",
    "AI_Advice",
    "AI_Evaluation",
    "AI_Caution",
]


class AIClient:
    """Wrapper around OpenAI chat API to produce strict JSON for task evaluation."""

    def __init__(self, model: str = "gpt-5", api_key: str | None = None, max_retries: int = 3, timeout: int = 60):
        # Load .env if present so users can place OPENAI_API_KEY there
        try:
            load_dotenv()
        except Exception:
            pass
        api_key = api_key or os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY が環境変数に設定されていません。")
        # The OpenAI SDK reads the key from env automatically as well, but set explicitly for clarity
        os.environ.setdefault("OPENAI_API_KEY", api_key)
        self.client = OpenAI()
        self.model = model
        self.max_retries = max_retries
        self.timeout = timeout

    def _build_messages(self, task: Dict[str, Any]) -> List[Dict[str, str]]:
        # System prompt enforces JSON-only response and field constraints
        system = (
            "あなたはプロジェクト管理の日本語アシスタントです。"  # noqa: E501
            "出力は必ず1つのJSONオブジェクトのみ、余計な文字や説明は禁止です。"
            "必ず以下のキーのみを含めてください: AI_工数, AI_Level, AI_NextAction, AI_Advice, AI_Evaluation, AI_Caution。"
            "制約: AI_Levelは『易』『並』『難』のいずれか。AI_Cautionは文字列配列。"
            "AI_工数は人日単位の概算（例: '0.5人日', '2人日'）。"
            "AI_工数は課題の工数見積もりを示してください。"
            "AI_Levelは、課題の難易度や複雑さに基づいて評価"
            "AI_NextActionは、現状把握と課題を顧みて、次にすべきアクションを箇条書きで記載してください。"
            "AI_Adviceは、現状把握と課題を顧みて、アドバイスを箇条書きで記載してください。"
            "AI_Evaluationは、No.、課題、ステータス、人員担当者、現状把握から、総合評価を箇条書きで記載してください。"
            "AI_Cautionは、現状把握と課題を顧みて、注意点を箇条書きで記載してください。"
        )
        user = (
            "以下の課題管理表の1行を評価し、上記のJSONキーで返答してください。\n"
            "可能なら現状把握や備考も考慮してください。\n\n"
            f"No.: {task.get('No.', '')}\n"
            f"課題: {task.get('課題', '')}\n"
            f"ステータス: {task.get('ステータス', '')}\n"
            f"人員: {task.get('人員', '')}\n"
            f"担当者: {task.get('担当者', '')}\n"
            f"現状把握: {task.get('現状把握', '')}\n"
            f"備考・コメント: {task.get('備考・コメント', '')}\n"
            "出力はJSONのみ。キーは指定の6つのみ。"
        )
        return [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]

    def _build_schema(self) -> Dict[str, Any]:
        return {
            "name": "ai_task_evaluation",
            "schema": {
                "type": "object",
                "properties": {
                    "AI_工数": {"type": "string"},
                    "AI_Level": {"type": "string", "enum": ["低", "中", "高"]},
                    "AI_NextAction": {"type": "string"},
                    "AI_Advice": {"type": "string"},
                    "AI_Evaluation": {"type": "string"},
                    "AI_Caution": {"type": "array", "items": {"type": "string"}},
                },
                "required": JSON_KEYS,
                "additionalProperties": False,
            },
            "strict": True,
        }

    def _messages_to_prompt(self, messages: List[Dict[str, str]]) -> str:
        # Flatten messages into a single prompt for Responses API fallback
        parts: List[str] = []
        for m in messages:
            role = m.get("role", "")
            content = m.get("content", "")
            if role == "system":
                parts.append(f"[SYSTEM]\n{content}")
            elif role == "user":
                parts.append(f"[USER]\n{content}")
            else:
                parts.append(content)
        return "\n\n".join(parts)

    def evaluate_task(self, task: Dict[str, Any]) -> Dict[str, Any]:
        messages = self._build_messages(task)

        last_error: Exception | None = None
        # Helper to detect temperature unsupported errors
        def _temp_unsupported(err: Exception) -> bool:
            s = str(err).lower()
            return ("temperature" in s) and ("unsupported" in s or "does not support" in s)

        for attempt in range(1, self.max_retries + 1):
            try:
                # Prefer JSON Schema when available (newer models/APIs)
                # First, attempt with temperature
                try:
                    resp = self.client.chat.completions.create(
                        model=self.model,
                        messages=messages,
                        temperature=0.2,
                        response_format={
                            "type": "json_schema",
                            "json_schema": self._build_schema(),
                        },
                    )
                except Exception as e0:
                    # If temperature unsupported, retry without temperature
                    if _temp_unsupported(e0):
                        try:
                            resp = self.client.chat.completions.create(
                                model=self.model,
                                messages=messages,
                                response_format={
                                    "type": "json_schema",
                                    "json_schema": self._build_schema(),
                                },
                            )
                        except Exception as e1:
                            # Fallback to simple json_object
                            try:
                                if _temp_unsupported(e1):
                                    resp = self.client.chat.completions.create(
                                        model=self.model,
                                        messages=messages,
                                        response_format={"type": "json_object"},
                                    )
                                else:
                                    resp = self.client.chat.completions.create(
                                        model=self.model,
                                        messages=messages,
                                        temperature=0.2,
                                        response_format={"type": "json_object"},
                                    )
                            except Exception as e2:
                                if _temp_unsupported(e2):
                                    resp = self.client.chat.completions.create(
                                        model=self.model,
                                        messages=messages,
                                        response_format={"type": "json_object"},
                                    )
                                else:
                                    raise
                    else:
                        # If JSON Schema unsupported or other error, try json_object with temp, then without temp
                        try:
                            resp = self.client.chat.completions.create(
                                model=self.model,
                                messages=messages,
                                temperature=0.2,
                                response_format={"type": "json_object"},
                            )
                        except Exception as e3:
                            if _temp_unsupported(e3):
                                resp = self.client.chat.completions.create(
                                    model=self.model,
                                    messages=messages,
                                    response_format={"type": "json_object"},
                                )
                            else:
                                raise

                content = resp.choices[0].message.content or "{}"
                data = json.loads(content)

                # Keep only expected keys, fill missing
                cleaned: Dict[str, Any] = {}
                for k in JSON_KEYS:
                    cleaned[k] = data.get(k, "")
                # Normalize AI_Caution to list[str]
                caut = cleaned.get("AI_Caution", [])
                if isinstance(caut, str):
                    # split by newline or bullet
                    parts = [p.strip(" ・-•\t") for p in caut.replace("\r", "").split("\n") if p.strip()]
                    cleaned["AI_Caution"] = parts
                elif isinstance(caut, list):
                    cleaned["AI_Caution"] = [str(x).strip() for x in caut if str(x).strip()]
                else:
                    cleaned["AI_Caution"] = []
                return cleaned
            except Exception as e:  # noqa: BLE001
                last_error = e
                # Try Responses API as a fallback (for future models like gpt-5)
                try:
                    prompt = self._messages_to_prompt(messages)
                    try:
                        resp2 = self.client.responses.create(
                            model=self.model,
                            input=prompt,
                            temperature=0.2,
                            response_format={
                                "type": "json_schema",
                                "json_schema": self._build_schema(),
                            },
                        )
                    except Exception as re0:
                        # Temperature unsupported? Retry without it
                        if _temp_unsupported(re0):
                            try:
                                resp2 = self.client.responses.create(
                                    model=self.model,
                                    input=prompt,
                                    response_format={
                                        "type": "json_schema",
                                        "json_schema": self._build_schema(),
                                    },
                                )
                            except Exception as re1:
                                # As last resort, no response_format
                                try:
                                    if _temp_unsupported(re1):
                                        resp2 = self.client.responses.create(
                                            model=self.model,
                                            input=prompt,
                                        )
                                    else:
                                        resp2 = self.client.responses.create(
                                            model=self.model,
                                            input=prompt,
                                            temperature=0.2,
                                        )
                                except Exception as re2:
                                    if _temp_unsupported(re2):
                                        resp2 = self.client.responses.create(
                                            model=self.model,
                                            input=prompt,
                                        )
                                    else:
                                        raise
                        else:
                            # json_schema or general failure: try without response_format
                            try:
                                resp2 = self.client.responses.create(
                                    model=self.model,
                                    input=prompt,
                                    temperature=0.2,
                                )
                            except Exception as re3:
                                if _temp_unsupported(re3):
                                    resp2 = self.client.responses.create(
                                        model=self.model,
                                        input=prompt,
                                    )
                                else:
                                    raise

                    # Newer SDKs expose aggregated text via output_text
                    content = getattr(resp2, "output_text", None)
                    if not content:
                        # Fallback: try to extract from outputs
                        content = ""
                        outputs = getattr(resp2, "output", None) or []
                        for block in outputs:
                            if isinstance(block, dict):
                                if block.get("type") == "output_text":
                                    content += str(block.get("text", ""))
                                elif block.get("type") == "message":
                                    content += str(block.get("content", ""))
                        if not content:
                            raise RuntimeError("No text content in Responses API output")

                    data = json.loads(content)

                    cleaned: Dict[str, Any] = {}
                    for k in JSON_KEYS:
                        cleaned[k] = data.get(k, "")
                    caut = cleaned.get("AI_Caution", [])
                    if isinstance(caut, str):
                        parts = [p.strip(" ・-•\t") for p in caut.replace("\r", "").split("\n") if p.strip()]
                        cleaned["AI_Caution"] = parts
                    elif isinstance(caut, list):
                        cleaned["AI_Caution"] = [str(x).strip() for x in caut if str(x).strip()]
                    else:
                        cleaned["AI_Caution"] = []
                    return cleaned
                except Exception as _:
                    # Backoff then retry next attempt
                    sleep_s = min(2 ** (attempt - 1), 8)
                    time.sleep(sleep_s)
                    continue

        # If we exit the retry loop
        raise RuntimeError(f"OpenAI呼び出しに失敗しました: {last_error}")
