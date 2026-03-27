import json
import ollama
from .prompt_schema import SYSTEM_PROMPT, EmailAnalysis


def _call_model(system_prompt: str, subject: str, body: str) -> str:
    """
    Gọi Llama 3.1 qua Ollama, ép output theo JSON schema của EmailAnalysis.
    Trả về chuỗi JSON thô.
    """
    user_prompt = f"Subject: {subject}\nBody: {body}"

    response = ollama.chat(
        model="llama3.1",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_prompt},
        ],
        format=EmailAnalysis.model_json_schema(),   
        options={"temperature": 0.1},              
    )

    return response["message"]["content"]


def process_row(row) -> tuple[str, str, str]:
    """
    Nhận một hàng (row) của DataFrame đã clean, trích xuất Subject + Body,
    gọi model AI và trả về tuple (category, urgency, summary).

    Parameters
    ----------
    row : pd.Series
        Hàng của DataFrame, cần có cột 'Subject' và 'Body'.

    Returns
    -------
    tuple[str, str, str]
        (category, urgency, summary)
        Trả về ("Error", "Error", "<thông báo lỗi>") nếu có ngoại lệ.
    """
    subject = str(row["Subject"]) if (row.notna()["Subject"]) else "Không có tiêu đề"
    body    = str(row["Body"])    if (row.notna()["Body"])    else "Không có nội dung"

    try:
        raw_json = _call_model(SYSTEM_PROMPT, subject, body)

        result = json.loads(raw_json)

        category = result.get("category", "Unknown")
        urgency  = result.get("urgency",  "Unknown")
        summary  = result.get("summary",  "Lỗi trích xuất")

        return category, urgency, summary

    except json.JSONDecodeError as e:
        print(f"[process_row] JSON parse error: {e} | raw='{raw_json}'")
        return "Error", "Error", "Invalid JSON from model"

    except Exception as e:
        print(f"[process_row] Unexpected error: {e}")
        return "Error", "Error", str(e)