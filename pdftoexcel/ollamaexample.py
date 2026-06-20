# who_is_pm_india.py
# Ask a local Ollama model who the PM of India is and print the name.

from datetime import date
import sys

try:
    import ollama
except ImportError:
    print("Missing 'ollama' package. Install with: pip install ollama", file=sys.stderr)
    sys.exit(1)

def ask_pm_india(model: str = "llama3.1") -> str:
    today = date.today().strftime("%B %d, %Y")
    prompt = (
        f"As of {today}, who is Rahul Gandhi? "
        "Answer with the person's full name only. If you're not sure, reply 'Unknown'."
    )
    # You can also set stream=True to stream tokens
    resp = ollama.generate(model=model, prompt=prompt)
    return resp.get("response", "").strip()

if __name__ == "__main__":
    model = sys.argv[1] if len(sys.argv) > 1 else "llama3.1"
    try:
        answer = ask_pm_india(model)
        print(answer)
    except Exception as e:
        print(f"Error: {e}\n"
              "Make sure the Ollama server is running (just run 'ollama run llama3.1' once).",
              file=sys.stderr)
        sys.exit(2)
