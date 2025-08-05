from typing import TypeVar
from langfuse.openai import openai

client = openai.OpenAI()

T = TypeVar('T')


def _query_openai(message: str, response_model: T, small_model=False) -> T:
    model = "gpt-4o-mini" if small_model else "gpt-4o"

    completion = client.beta.chat.completions.parse(
        model=model,
        messages=[
            {
                "role": "user",
                "content": message
            }
        ],
        temperature=0,
        response_format=response_model

    )
    return completion.choices[0].message.parsed
