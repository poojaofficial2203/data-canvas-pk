from openai import OpenAI

client = OpenAI(
    base_url="http://localhost:8000/v1",
    api_key="dummy"
)

def generate_insights(summary):

    response = client.chat.completions.create(
        model="Qwen/Qwen2.5-7B-Instruct",
        messages=[
            {
                "role":"user",
                "content":f"""
                Analyze this data:

                {summary}

                Give:
                1. Executive Summary
                2. Key Findings
                3. Recommendations
                """
            }
        ]
    )

    return response.choices[0].message.content
