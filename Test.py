from ddgs import DDGS
from openai import OpenAI
import os

# 1. Test DuckDuckGo
print("🔎 Testing DuckDuckGo search...")
with DDGS() as ddgs:
    results = [r["body"] for r in ddgs.text("latest AI news", max_results=3)]
print("DuckDuckGo results:", results, "\n")

# 2. Test OpenAI
print("🤖 Testing OpenAI API...")
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

resp = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[{"role": "user", "content": "Say hello from my internship setup script!"}],
)

print("OpenAI says:", resp.choices[0].message.content)
