"""
Auto Insurance Data Q&A Chatbot
Flask app + Claude / Gemini API for natural language data queries.
"""
import json, os
from flask import Flask, request, jsonify, render_template_string
import anthropic

try:
    from google import genai as google_genai
    GOOGLE_AI_AVAILABLE = True
except ImportError:
    GOOGLE_AI_AVAILABLE = False

# ── Load and pre-aggregate data ──
with open(os.path.join(os.path.dirname(__file__), "dashboard_data.json")) as f:
    raw = json.load(f)

records = []
for r in raw:
    records.append({
        "id": r[0], "gender": "Male" if r[1] == 1 else "Female",
        "age": r[2], "car_year": r[3], "premium": r[4], "loss": r[5],
        "loss_ratio": round(r[5] / r[4], 4) if r[4] > 0 else 0
    })

def age_group(a):
    if a < 25: return "18-24"
    if a < 35: return "25-34"
    if a < 45: return "35-44"
    if a < 55: return "45-54"
    if a < 65: return "55-64"
    return "65+"

def car_era(y):
    if y < 2005: return "2000-2004"
    if y < 2010: return "2005-2009"
    if y < 2015: return "2010-2014"
    if y < 2020: return "2015-2019"
    return "2020-2025"

# Pre-compute summary for context
total = len(records)
total_prem = sum(r["premium"] for r in records)
total_loss = sum(r["loss"] for r in records)
claims = [r for r in records if r["loss"] > 0]
loss_ratio = total_loss / total_prem

age_groups = ["18-24", "25-34", "35-44", "45-54", "55-64", "65+"]
by_age = {}
for ag in age_groups:
    rows = [r for r in records if age_group(r["age"]) == ag]
    p = sum(r["premium"] for r in rows)
    l = sum(r["loss"] for r in rows)
    c = sum(1 for r in rows if r["loss"] > 0)
    by_age[ag] = {"count": len(rows), "avg_premium": round(p/len(rows), 2),
                  "loss_ratio": round(l/p*100, 1) if p else 0,
                  "claims_freq": round(c/len(rows)*100, 1) if rows else 0,
                  "total_premium": round(p, 2), "total_loss": round(l, 2)}

by_gender = {}
for g in ["Male", "Female"]:
    rows = [r for r in records if r["gender"] == g]
    p = sum(r["premium"] for r in rows)
    l = sum(r["loss"] for r in rows)
    c = sum(1 for r in rows if r["loss"] > 0)
    by_gender[g] = {"count": len(rows), "avg_premium": round(p/len(rows), 2),
                    "loss_ratio": round(l/p*100, 1), "claims_freq": round(c/len(rows)*100, 1)}

eras = ["2000-2004", "2005-2009", "2010-2014", "2015-2019", "2020-2025"]
by_era = {}
for e in eras:
    rows = [r for r in records if car_era(r["car_year"]) == e]
    p = sum(r["premium"] for r in rows)
    l = sum(r["loss"] for r in rows)
    by_era[e] = {"count": len(rows), "loss_ratio": round(l/p*100, 1) if p else 0}

top_claims = sorted(records, key=lambda r: r["loss"], reverse=True)[:20]
top_claims_summary = json.dumps([
    {"id": c["id"], "gender": c["gender"], "age": c["age"], "car_year": c["car_year"],
     "premium": c["premium"], "loss": c["loss"], "loss_ratio": round(c["loss"]/c["premium"], 1)}
    for c in top_claims[:10]
], indent=2)

DATA_CONTEXT = f"""You are an AI data analyst for an auto insurance company. You answer questions about a portfolio of {total:,} policyholders.

KEY METRICS:
- Total policies: {total:,}
- Total annual premium collected: ${total_prem:,.0f}
- Total claims paid: ${total_loss:,.0f}
- Overall loss ratio: {loss_ratio*100:.1f}%
- Claims frequency: {len(claims)/total*100:.1f}% (i.e., {len(claims):,} out of {total:,} had claims)
- Average premium: ${total_prem/total:,.0f}
- Average claim (claimants only): ${total_loss/len(claims):,.0f}

BY AGE GROUP:
{json.dumps(by_age, indent=2)}

BY GENDER:
{json.dumps(by_gender, indent=2)}

BY CAR ERA:
{json.dumps(by_era, indent=2)}

TOP 10 LARGEST CLAIMS:
{top_claims_summary}

RULES:
- Answer concisely and precisely using the data above
- Use specific numbers from the data, not approximations
- If asked something not in the data, say so clearly
- Format currency with $ and commas
- Format percentages with one decimal place
- Keep responses brief (2-4 sentences) unless the user asks for detail
"""

app = Flask(__name__)

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Insurance Data Q&A — AI Chatbot</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{--bg:#0f1117;--surface:#1a1d27;--surface2:#232733;--border:#2e3345;--text:#e4e6ed;--muted:#8b8fa3;--accent:#6366f1;--accent-light:#818cf8;--cyan:#06b6d4;--green:#22c55e}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg);color:var(--text);height:100vh;display:flex;flex-direction:column}
header{background:var(--surface);border-bottom:1px solid var(--border);padding:16px 24px;display:flex;align-items:center;justify-content:space-between}
header h1{font-size:1.2rem;font-weight:700;color:var(--accent-light)}
header .subtitle{font-size:.8rem;color:var(--muted)}
.api-bar{background:var(--surface2);border-bottom:1px solid var(--border);padding:10px 24px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.api-bar label{font-size:.75rem;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
.api-bar input{flex:1;max-width:500px;background:var(--surface);color:var(--text);border:1px solid var(--border);padding:8px 12px;border-radius:6px;font-size:.85rem;font-family:monospace}
.api-bar input:focus{outline:none;border-color:var(--accent)}
.api-bar .status{font-size:.75rem;padding:4px 10px;border-radius:12px}
.status-ready{background:rgba(34,197,94,.15);color:var(--green)}
.status-missing{background:rgba(239,68,68,.15);color:#ef4444}
.model-select{background:var(--surface);color:var(--text);border:1px solid var(--border);padding:8px 12px;border-radius:6px;font-size:.85rem;cursor:pointer;outline:none}
.model-select:focus{border-color:var(--accent)}
.model-select option{background:var(--surface);color:var(--text)}
.chat-area{flex:1;overflow-y:auto;padding:24px;display:flex;flex-direction:column;gap:16px}
.msg{max-width:75%;padding:14px 18px;border-radius:12px;font-size:.9rem;line-height:1.6;animation:fadeIn .3s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
.msg-user{align-self:flex-end;background:var(--accent);color:white;border-bottom-right-radius:4px}
.msg-bot{align-self:flex-start;background:var(--surface);border:1px solid var(--border);border-bottom-left-radius:4px}
.msg-bot .label{font-size:.7rem;color:var(--cyan);text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px}
.msg-system{align-self:center;background:var(--surface2);color:var(--muted);font-size:.8rem;border-radius:20px;padding:8px 20px}
.typing{align-self:flex-start;padding:14px 18px;background:var(--surface);border:1px solid var(--border);border-radius:12px}
.typing span{display:inline-block;width:8px;height:8px;background:var(--muted);border-radius:50%;margin:0 2px;animation:bounce .6s infinite alternate}
.typing span:nth-child(2){animation-delay:.2s}
.typing span:nth-child(3){animation-delay:.4s}
@keyframes bounce{to{transform:translateY(-6px);opacity:.3}}
.input-area{background:var(--surface);border-top:1px solid var(--border);padding:16px 24px;display:flex;gap:12px}
.input-area input{flex:1;background:var(--surface2);color:var(--text);border:1px solid var(--border);padding:12px 16px;border-radius:8px;font-size:.9rem}
.input-area input:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 2px rgba(99,102,241,.2)}
.input-area button{background:var(--accent);color:white;border:none;padding:12px 24px;border-radius:8px;font-size:.9rem;font-weight:600;cursor:pointer;transition:background .2s}
.input-area button:hover{background:var(--accent-light)}
.input-area button:disabled{opacity:.5;cursor:not-allowed}
.suggestions{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}
.suggestions button{background:var(--surface2);color:var(--muted);border:1px solid var(--border);padding:6px 14px;border-radius:16px;font-size:.78rem;cursor:pointer;transition:all .2s}
.suggestions button:hover{border-color:var(--accent);color:var(--accent-light)}
</style>
</head>
<body>
<header>
  <div>
    <h1>Insurance Data Q&A</h1>
    <div class="subtitle">Ask questions about 10,000 policyholders in plain English</div>
  </div>
  <div style="text-align:right">
    <div style="font-size:.7rem;color:var(--muted)" id="headerPowered">Powered by Claude AI</div>
  </div>
</header>

<div class="api-bar">
  <label id="apiKeyLabel">API Key</label>
  <select class="model-select" id="modelSelect" onchange="onModelChange()">
    <optgroup label="Anthropic">
      <option value="claude-sonnet-4-6" data-provider="anthropic">Claude claude-sonnet-4-6 (Anthropic)</option>
      <option value="claude-haiku-4-5-20251001" data-provider="anthropic">Claude claude-haiku-4-5-20251001 (Anthropic)</option>
    </optgroup>
    <optgroup label="Google">
      <option value="gemini-2.0-flash" data-provider="google">Gemini 2.0 Flash (Google)</option>
      <option value="gemini-1.5-pro" data-provider="google">Gemini 1.5 Pro (Google)</option>
      <option value="gemini-3-flash-preview" data-provider="google">Gemini 3 Flash Preview (Google)</option>
    </optgroup>
  </select>
  <input type="password" id="apiKey" placeholder="sk-ant-..." autocomplete="off">
  <span class="status status-missing" id="apiStatus">Enter API key</span>
</div>

<div class="chat-area" id="chat">
  <div class="msg-system">Welcome! Enter your API key above, then ask me anything about the insurance portfolio.</div>
  <div class="suggestions" id="suggestions">
    <button onclick="askSuggestion(this)">Which age group has the highest loss ratio?</button>
    <button onclick="askSuggestion(this)">Compare male vs female claims frequency</button>
    <button onclick="askSuggestion(this)">What's the riskiest segment?</button>
    <button onclick="askSuggestion(this)">Show me the top 5 largest claims</button>
    <button onclick="askSuggestion(this)">How do older cars affect loss ratios?</button>
  </div>
</div>

<div class="input-area">
  <input type="text" id="userInput" placeholder="Ask about the data... (e.g., 'What is the average premium for young drivers?')" disabled>
  <button id="sendBtn" onclick="sendMessage()" disabled>Send</button>
</div>

<script>
const chatEl = document.getElementById('chat');
const inputEl = document.getElementById('userInput');
const sendBtn = document.getElementById('sendBtn');
const apiKeyEl = document.getElementById('apiKey');
const apiStatus = document.getElementById('apiStatus');
const modelSelect = document.getElementById('modelSelect');
const apiKeyLabel = document.getElementById('apiKeyLabel');
const headerPowered = document.getElementById('headerPowered');

function getSelectedProvider() {
  const selectedOption = modelSelect.options[modelSelect.selectedIndex];
  return selectedOption.getAttribute('data-provider');
}

function onModelChange() {
  const provider = getSelectedProvider();
  if (provider === 'google') {
    apiKeyEl.placeholder = 'AIza...';
    apiKeyLabel.textContent = 'Google API Key';
    headerPowered.textContent = 'Powered by Google Gemini';
  } else {
    apiKeyEl.placeholder = 'sk-ant-...';
    apiKeyLabel.textContent = 'API Key';
    headerPowered.textContent = 'Powered by Claude AI';
  }
  // Re-check key validity after switching
  const hasKey = apiKeyEl.value.trim().length > 10;
  inputEl.disabled = !hasKey;
  sendBtn.disabled = !hasKey;
  apiStatus.textContent = hasKey ? 'Ready' : 'Enter API key';
  apiStatus.className = 'status ' + (hasKey ? 'status-ready' : 'status-missing');
}

apiKeyEl.addEventListener('input', () => {
  const hasKey = apiKeyEl.value.trim().length > 10;
  inputEl.disabled = !hasKey;
  sendBtn.disabled = !hasKey;
  apiStatus.textContent = hasKey ? 'Ready' : 'Enter API key';
  apiStatus.className = 'status ' + (hasKey ? 'status-ready' : 'status-missing');
});

inputEl.addEventListener('keydown', e => { if (e.key === 'Enter' && !sendBtn.disabled) sendMessage(); });

function addMsg(text, type) {
  const div = document.createElement('div');
  if (type === 'user') {
    div.className = 'msg msg-user';
    div.textContent = text;
  } else if (type === 'bot') {
    div.className = 'msg msg-bot';
    div.innerHTML = '<div class="label">AI Analyst</div>' + text.replace(/\n/g, '<br>');
  } else {
    div.className = 'msg-system';
    div.textContent = text;
  }
  chatEl.appendChild(div);
  chatEl.scrollTop = chatEl.scrollHeight;
  return div;
}

function showTyping() {
  const div = document.createElement('div');
  div.className = 'typing';
  div.id = 'typing';
  div.innerHTML = '<span></span><span></span><span></span>';
  chatEl.appendChild(div);
  chatEl.scrollTop = chatEl.scrollHeight;
}
function hideTyping() { const t = document.getElementById('typing'); if (t) t.remove(); }

async function sendMessage() {
  const text = inputEl.value.trim();
  if (!text) return;
  inputEl.value = '';

  // Hide suggestions after first message
  const sug = document.getElementById('suggestions');
  if (sug) sug.remove();

  addMsg(text, 'user');
  sendBtn.disabled = true;
  inputEl.disabled = true;
  showTyping();

  const selectedOption = modelSelect.options[modelSelect.selectedIndex];
  const model = modelSelect.value;
  const provider = selectedOption.getAttribute('data-provider');

  try {
    const res = await fetch('/ask', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        question: text,
        api_key: apiKeyEl.value.trim(),
        model: model,
        provider: provider
      })
    });
    hideTyping();
    const data = await res.json();
    if (data.error) {
      addMsg('Error: ' + data.error, 'system');
    } else {
      addMsg(data.answer, 'bot');
    }
  } catch (err) {
    hideTyping();
    addMsg('Connection error: ' + err.message, 'system');
  }

  sendBtn.disabled = false;
  inputEl.disabled = false;
  inputEl.focus();
}

function askSuggestion(btn) {
  inputEl.value = btn.textContent;
  sendMessage();
}
</script>
</body>
</html>"""

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/ask", methods=["POST"])
def ask():
    body = request.json
    question = body.get("question", "").strip()
    api_key = body.get("api_key", "").strip()
    model = body.get("model", "claude-sonnet-4-6").strip()
    provider = body.get("provider", "anthropic").strip()

    if not api_key:
        return jsonify({"error": "API key is required"}), 400
    if not question:
        return jsonify({"error": "Question is required"}), 400

    if provider == "google":
        if not GOOGLE_AI_AVAILABLE:
            return jsonify({"error": "google-genai package is not installed on the server."}), 500
        try:
            client = google_genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model=model,
                contents=question,
                config={"system_instruction": DATA_CONTEXT}
            )
            answer = response.text
            return jsonify({"answer": answer})
        except Exception as e:
            err_str = str(e)
            if "API_KEY_INVALID" in err_str or "invalid" in err_str.lower() or "api key" in err_str.lower():
                return jsonify({"error": "Invalid Google API key. Please check and try again."}), 401
            return jsonify({"error": err_str}), 500
    else:
        # Default: Anthropic
        try:
            client = anthropic.Anthropic(api_key=api_key)
            message = client.messages.create(
                model=model,
                max_tokens=500,
                system=DATA_CONTEXT,
                messages=[{"role": "user", "content": question}]
            )
            answer = message.content[0].text
            return jsonify({"answer": answer})
        except anthropic.AuthenticationError:
            return jsonify({"error": "Invalid API key. Please check and try again."}), 401
        except Exception as e:
            return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  Insurance Data Q&A Chatbot")
    print("  Open in browser: http://localhost:5000")
    print("=" * 60 + "\n")
    app.run(debug=False, port=5000)
