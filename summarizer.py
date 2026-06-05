import anthropic


def generate_summary(risk_table_text: str, counts: dict[str, int], api_key: str) -> str:
    """Call Claude API and return Korean-language threat narrative."""
    if not api_key:
        return "(API 키가 설정되지 않아 AI 요약을 생성하지 못했습니다.)"

    counts_text = "\n".join(f"- {k}: {v}건" for k, v in counts.items())
    prompt = f"""당신은 회사 보안 담당자의 DLP 분석을 돕는 보안 전문가입니다.
아래는 이번 달 DLP 탐지 결과입니다.

[탐지 건수 요약]
{counts_text}

[상위 위험 사용자]
{risk_table_text}

다음 세 가지를 한국어로 작성하세요:
1. 이달의 위협 총평 (2~3문장)
2. 상위 3명 사용자의 행동 패턴 코멘트 (각 1~2문장)
3. 즉시 조치 권고사항 (불릿 포인트 3~5개)

간결하고 실무적으로 작성하세요."""

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text
