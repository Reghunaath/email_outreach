// Clicking the toolbar icon opens (and toggles) the side panel instead of a
// dropdown popup, so it stays open until dismissed.
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch(() => {});

const DEFAULT_MODEL = "gemini-flash-latest";
const ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models";
const ASHBY_POSTING_API = "https://api.ashbyhq.com/posting-api/job-board";
const BOARD_TTL = 5 * 60 * 1000; // cache a board's postings for 5 minutes

const boardCache = new Map(); // board name -> { ts, jobs }

// The JD is fetched from Ashby's public posting API and matched by posting ID.
// board + postingId come from the page URL: /{board}/{postingId}[/application].
async function fetchJobDescription(board, postingId) {
  if (!board || !postingId) return "";
  try {
    let entry = boardCache.get(board);
    if (!entry || Date.now() - entry.ts > BOARD_TTL) {
      const r = await fetch(
        `${ASHBY_POSTING_API}/${encodeURIComponent(board)}?includeCompensation=true`,
      );
      if (!r.ok) return "";
      const d = await r.json();
      entry = { ts: Date.now(), jobs: Array.isArray(d.jobs) ? d.jobs : [] };
      boardCache.set(board, entry);
    }
    const job = entry.jobs.find((j) => j.id === postingId);
    return job ? job.descriptionPlain || "" : "";
  } catch (_) {
    return "";
  }
}

// Hard-coded applicant profile (from Reghunaath_Resume_May_N.pdf). Anything the
// user types into the options "profile" field is appended to this at prompt time.
const PROFILE_BASE = `Name: Reghunaath Ajith Kumar Ahila
Email: reghunaath4@gmail.com
Phone: (857) 351-9009
Portfolio: https://reghunaath.com/
Location: Boston, MA (open to relocation)

EDUCATION
- Northeastern University, Boston, MA, USA — Master of Science in Data Science (Sep 2024 – Apr 2026), CGPA 3.9.
- Vellore Institute of Technology, Vellore, TN, India — Bachelor of Technology in Computer Science and Engineering (Aug 2018 – Apr 2022), CGPA 8.56.

WORK EXPERIENCE
QuantUniversity, Boston, MA — Graduate Intern (Jul 2025 – Aug 2025)
- Built a platform using React and FastAPI to enable AI-assisted educational content creation and seamless hosting of generated materials, reducing content development time from 5 days to ~3 hours.
- Identified and fixed a critical paywall bypass vulnerability in the first week, strengthening platform security.
- Designed and implemented authentication and authorization systems addressing the security and compliance requirements for ISO and SOC2 certification.

Infosys, Bengaluru, India — Senior Systems Engineer (Aug 2023 – Jul 2024)
- Developed and deployed a full-stack application with a .NET microservice architecture and React.js frontend, modernizing a legacy insurance platform through REST APIs, JWT-based authentication, and Redux state management.
- Built a .NET rule-based recommendation engine integrating 9 external systems through gRPC, SOAP, and REST APIs, with SQL used to cache external system data for efficient rule evaluation and policy recommendations.
- Developed a Python script to auto-generate unit test cases and Postman integration test cases from business-owned Excel sheets, reducing test-case updates from hours to minutes per cycle and saving over 65 hours of manual effort long term.

Danske IT (subsidiary of Danske Bank, acquired by Infosys in 2023), Bengaluru, India — Associate Software Engineer (Jul 2022 – Aug 2023)
- Developed and integrated Camunda BPM workflows within the .NET backend to orchestrate customer onboarding processes, automating task execution and service interactions to improve data processing efficiency, fault tolerance, and overall system reliability.
- Built and owned CI/CD pipelines on Azure DevOps, ensuring smooth deployment workflows, version control, and continuous integration across development and production environments.
- Integrated automated load testing with Grafana K6 into the CD pipeline to evaluate system performance and ensure scalability under high traffic.
- Independently implemented a monitoring solution using Kibana (Elastic Stack) to provide real-time insights across multiple team projects.

Danske IT, Bengaluru, India — Apprentice (Jan 2022 – Jul 2022)
- Gained comprehensive experience in fintech software development, working across testing, DevOps, frontend, and backend in an agile environment.
- Improved unit test line coverage from 60% to 95% for the .NET backend.

PROJECTS
- RescueLine AI — AI-powered emergency call triage system using Twilio, ElevenLabs, and FastAPI to automatically classify and route emergency calls by urgency during natural disasters when traditional helplines are overwhelmed. Built a real-time voice AI agent and live dashboard with WebSocket-based updates for emergency coordinators managing high call volumes. Earned 1st place and $700 at Northeastern's Innovate 2026 Hackathon.
- LeadCatch AI — AI chat assistant powered by ChatGPT and Twilio APIs to turn missed calls into booked appointments for small businesses. Designed a scalable Python backend for multi-user handling and automated SMS-based lead conversion. Earned 2nd place and $1500 in the "Best Startup Demonstrating Traction" category at the Yconic AI Hackathon (Microsoft, Cambridge, MA).
- Doodlpop (1st Place, SharkHack Hackathon) — AI-powered comic book generator that turns a single sentence into a fully illustrated comic. From a story idea it generates a panel-by-panel script with dialogue and visual descriptions, lets the user pick an art style (manga, western, watercolor storybook), edit the script, then illustrates every panel with AI. Supports shareable links, QR code sharing, and PDF export.
- SNAPBACK (2nd Place, Babson Generator Build-a-thon) — computer vision tool that measures athletic mobility loss after injury or a long break with no wearables and no clinic visit. Uses MediaPipe and OpenCV to track 33 skeletal landmarks at 30fps and compute joint angles in real time, producing a mobility score out of 100 benchmarked against clinical reference ranges, a sport-specific gap analysis, and a personalised week-by-week return-to-sport exercise plan with sets, reps, and reasoning.

TECHNICAL SKILLS
- Backend: C# (.NET), Java (Spring Boot), Python (Django, Flask, FastAPI)
- Frontend: React, Angular, Flutter
- DevOps: AWS, Azure DevOps, git
- Other programming languages: C, C++, MATLAB, R, Go
- Testing: JUnit, NUnit, xUnit, Postman, Selenium, k6
- ML/AI: TensorFlow, Keras, PyTorch, scikit-learn, Pandas, NumPy, OpenCV, Matplotlib, Seaborn, Plotly, LLMs, Streamlit, GenAI, LangChain, LangGraph, ElevenLabs, RAG, Claude Code, Cursor
- Databases: SQL, MongoDB

PUBLICATIONS
- Muralidharan, K., Ramesh, A., Rithvik, G., Prem, S., Reghunaath, A. A., & Gopinath, M. P. (2021). "1D Convolution approach to human activity recognition using sensor data and comparison with machine learning algorithms." International Journal of Cognitive Computing in Engineering, 2, 130-143. (63 citations)

POSITIONING (lean on this for "why are you a good fit", "why you", or "why this role" style questions)
- Across four hackathons I have won top placements (two first places and two second places), with over $2,000 in combined prizes. Use this as evidence of strong product sense and ownership: I take a rough idea, decide what is actually worth building, and ship a working product end to end under tight time pressure, then carry it through to sharing and polish. Tie that same instinct to what the role and company need.`;

// Static persona + writing rules. Identical on every request, so it is sent as
// the model's systemInstruction rather than mixed into the per-request contents.
const SYSTEM_RULES = [
  "<role>",
  "You are helping a job applicant fill out an application form.",
  "Write a first-person answer to the application question, drawing on the",
  "applicant's profile and the job description.",
  "</role>",
  "",
  "<rules>",
  "- Write like a real person casually explaining themselves. Use a relaxed,",
  "  conversational tone and plain, everyday words, the way you would actually",
  "  talk, while still being appropriate for a job application.",
  '- Avoid fancy, formal, or buzzword-y words (for example "thrive", "bridge",',
  '  "leverage", "passionate", "highly polished", "seamless", "spearhead").',
  "  Keep technical terms exactly as they are. It must not sound AI-generated",
  "  or robotic.",
  '- Do NOT use em dashes (the "—" character) anywhere. Use commas, periods,',
  "  parentheses, or separate sentences instead.",
  "- Do NOT use hyphens to join words. Split the words with a space instead",
  '  (for example write "full stack" not "full-stack", "real time" not "real-time").',
  '- Do NOT use colons (the ":" character) anywhere. Rephrase into full sentences instead.',
  "- Keep the answer to about 100 words, unless the question explicitly asks for a",
  "  different length.",
  "- Return ONLY the answer text, ready to paste directly into the field: no preamble,",
  "  no surrounding quotes, and no markdown.",
  "- Be specific and ground the answer in the profile and job description.",
  "- For more specific questions, when the question relates to a skill, domain, or",
  "  experience that one of the applicant's projects or hackathon wins demonstrates,",
  "  reference that specific project by name and briefly what was built or achieved.",
  "  Prefer a concrete example over a generic claim.",
  "- Whenever you mention a specific hackathon project, also state the placement it",
  "  won and the prize money, using the details in the profile (for example RescueLine",
  "  AI won 1st place and $700, LeadCatch AI won 2nd place and $1500). If the profile",
  "  lists no prize money for that project, just state the placement.",
  "- Whenever you mention hackathon wins in general (without naming a specific",
  "  hackathon), state that the applicant has won 4 hackathons with over $2000 in",
  "  combined prize money.",
  "- If the job description is '(not available)', rely entirely on the applicant profile to answer the question.",
  "- If the question asks about the hardest, toughest, most difficult, or most",
  "  challenging problem, technical challenge, or obstacle the applicant has solved",
  "  or overcome, return the following answer verbatim, exactly as written, ignoring",
  "  the ~100 word length cap and all other formatting rules for this one answer:",
  '  "The hardest problem I\'ve solved was keeping characters visually consistent',
  "  across pages in DoodlPop, an AI comic book generator. Image models happily",
  "  redraw the same character as a different person on every page, so I attacked it",
  "  from a few angles at once. I wrote richer, more specific character descriptions",
  "  in the prompts, generated a character reference sheet that gets passed into every",
  "  page, fed the previous page back in as a reference so each new image stays",
  "  anchored to the last, and where it fit, based a character on a well known one the",
  "  model already understood. No single trick was enough. Stacking them was what",
  '  finally made a character feel like the same person from the first page to the last."',
  "</rules>",
].join("\n");

// Per-request inputs: the variable data the model grounds its answer in. Kept in
// contents (data first, question last) rather than in the system instruction.
function buildUserContent(question, jobDescription, profile, remark) {
  const profileSection = profile
    ? `${PROFILE_BASE}\n\nADDITIONAL NOTES FROM THE APPLICANT:\n${profile}`
    : PROFILE_BASE;
  const lines = [
    "<applicant_profile>",
    profileSection,
    "</applicant_profile>",
    "",
    "<job_description>",
    jobDescription || "(not available)",
    "</job_description>",
    "",
    "<application_question>",
    question,
    "</application_question>",
  ];
  if (remark) {
    lines.push(
      "",
      "<extra_instruction>",
      "Extra context or instruction from the applicant for this specific answer. Follow it.",
      remark,
      "</extra_instruction>",
    );
  }
  return lines.join("\n");
}

function extractText(data) {
  const parts = data?.candidates?.[0]?.content?.parts;
  if (!Array.isArray(parts)) return "";
  // Thinking models (e.g. gemini-3.x flash) may return multiple parts; keep the answer text.
  return parts
    .filter((p) => typeof p.text === "string")
    .map((p) => p.text)
    .join("")
    .trim();
}

async function generate({ question, board, postingId, pageText, remark }) {
  const {
    apiKey = "",
    model = "",
    profile = "",
  } = await chrome.storage.local.get(["apiKey", "model", "profile"]);

  if (!apiKey) {
    return {
      error:
        "No API key set. Open the extension options and add your Gemini API key.",
    };
  }
  if (!question) {
    return { error: "Could not determine the question for this field." };
  }

  let jobDescription = await fetchJobDescription(board, postingId);
  if (!jobDescription) jobDescription = pageText || "";

  const usedModel = model || DEFAULT_MODEL;
  const url = `${ENDPOINT}/${encodeURIComponent(usedModel)}:generateContent`;
  const body = {
    systemInstruction: { parts: [{ text: SYSTEM_RULES }] },
    contents: [
      {
        parts: [
          { text: buildUserContent(question, jobDescription, profile, remark) },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.4,
      topP: 0.95,
      responseMimeType: "text/plain",
    },
  };

  let res;
  try {
    res = await fetch(url, {
      method: "POST",
      headers: { "x-goog-api-key": apiKey, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
  } catch (err) {
    return { error: "Network error reaching Gemini: " + err.message };
  }

  let data;
  try {
    data = await res.json();
  } catch (_) {
    return {
      error: `Gemini returned an unreadable response (HTTP ${res.status}).`,
    };
  }

  if (!res.ok) {
    const message = data?.error?.message || `HTTP ${res.status}`;
    return { error: "Gemini error: " + message };
  }

  const text = extractText(data);
  if (!text) {
    return { error: "Gemini returned no answer text." };
  }
  return { text };
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg?.type === "ASHBY_FILL") {
    generate(msg).then(sendResponse);
    return true; // keep the message channel open for the async response
  }
});
