$ErrorActionPreference = 'Stop'

$sourceDocx = (Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*U1-U10*.docx' |
  Where-Object { -not $_.Name.StartsWith('~$') } |
  Select-Object -First 1 -ExpandProperty FullName)
$outputHtml = Join-Path $PSScriptRoot 'U1-U10_Difficult_Words_Interactive_Quiz.html'

Add-Type -AssemblyName System.IO.Compression
$stream = [IO.File]::Open($sourceDocx, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
$archive = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read)
$entry = $archive.GetEntry('word/document.xml')
$reader = New-Object IO.StreamReader($entry.Open())
[xml]$xml = $reader.ReadToEnd()
$reader.Close()
$archive.Dispose()
$stream.Dispose()

$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$lines = @(
  $xml.SelectNodes('//w:p', $ns) |
    ForEach-Object {
      ($_.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.'#text' }) -join ''
    } |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ }
)

function Get-Category([string]$rawPos) {
  $pos = $rawPos.ToLowerInvariant()
  if ($pos -match 'adj') { return 'adj' }
  if ($pos -match 'adv') { return 'adv' }
  if ($pos -match 'v') { return 'v' }
  if ($pos -match 'n') { return 'n' }
  return 'phr'
}

function Test-PosLine([string]$line) {
  return $line -match '^\(?\s*(n\.?|v\.?|adj\.?|adv\.?|phrase|phr\.?v|v\.?phr|adj\.?phr|n\.?phr|v/n)\s*\)?$'
}

$items = New-Object System.Collections.Generic.List[object]
$seen = @{}
for ($i = 1; $i -lt ($lines.Count - 4); $i++) {
  if (-not (Test-PosLine $lines[$i])) { continue }
  $word = $lines[$i - 1]
  $pos = $lines[$i].Trim('(', ')')
  $def = $lines[$i + 1]
  $cn = $lines[$i + 2]
  $sentence = $lines[$i + 3]
  $cnSentence = $lines[$i + 4]
  if ($word.Length -gt 80 -or $word -match '4000 essential') { continue }
  if ($sentence -notmatch '_{3,}') { continue }
  $key = $word.ToLowerInvariant()
  if ($seen.ContainsKey($key)) { continue }
  $seen[$key] = $true
  $items.Add([ordered]@{
    word = $word
    pos = $pos
    cat = Get-Category $pos
    en = $def.TrimEnd(':')
    cn = $cn
    sentence = $sentence
    cnSentence = $cnSentence
  })
}

$testableCount = @(
  $items | Where-Object {
    $cat = $_.cat
    $targetWord = $_.word
    @($items | Where-Object { $_.cat -eq $cat -and $_.word -ne $targetWord }).Count -ge 3
  }
).Count

if ($testableCount -lt 4) {
  throw "Not enough quiz entries were extracted from the Word file."
}

$json = $items | ConvertTo-Json -Depth 4 -Compress
$generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

$html = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>4000 Essential Words Level 5 | Units 1-10 Difficult Words Quiz</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{--primary:#1a3a5c;--primary2:#2a5298;--accent:#c8962e;--ok:#2d7a4f;--bad:#b53a2f;--bg:#faf8f5;--card:#fff;--line:#dde3ea;--muted:#6b7a8d;--clue:#fff3a3}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:"DM Sans",sans-serif;background:var(--bg);color:#1e2a35;min-height:100vh}
header{position:sticky;top:0;z-index:10;background:var(--primary);color:#fff;padding:14px 18px;display:flex;justify-content:space-between;align-items:center;gap:12px;box-shadow:0 2px 10px #0003}
.title{font-family:"Playfair Display",serif;font-size:1.05rem}.title span{color:#f1c86b}.score{white-space:nowrap;background:#ffffff22;border-radius:20px;padding:5px 13px;font-size:.84rem;font-weight:700}
.wrap{max-width:780px;margin:0 auto;padding:18px 14px 56px}.card{background:var(--card);border:1px solid var(--line);border-radius:14px;box-shadow:0 3px 16px #1a3a5c14;padding:20px;margin-bottom:14px}
h1{font-family:"Playfair Display",serif;color:var(--primary);font-size:1.75rem;margin-bottom:9px}.sub{font-size:.91rem;line-height:1.65;color:var(--muted);margin-bottom:15px}
label{display:block;font-size:.77rem;font-weight:700;letter-spacing:.05em;color:var(--muted);text-transform:uppercase;margin:11px 0 6px}
input,select{width:100%;padding:11px 12px;border:1.5px solid var(--line);border-radius:9px;background:#fff;font:inherit}input:focus,select:focus{outline:none;border-color:var(--primary)}
button{font:inherit;cursor:pointer;border:0}.primary{width:100%;margin-top:16px;padding:13px;border-radius:10px;background:var(--primary);color:#fff;font-weight:700}.primary:hover{background:var(--primary2)}
.progress-row{display:flex;justify-content:space-between;font-size:.8rem;color:var(--muted);font-weight:700;margin-bottom:7px}.track{height:6px;background:#dde3ea;border-radius:4px;overflow:hidden;margin-bottom:14px}.fill{height:100%;background:var(--accent);transition:width .3s}
.qnum{font-size:.76rem;color:var(--accent);font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin-bottom:8px}.passage{font-size:1rem;line-height:1.8;margin-bottom:14px}.blank{font-weight:700;color:var(--primary);border-bottom:2px solid var(--accent)}
.w{cursor:pointer;border-bottom:1px dotted #aaa}.w:hover{background:#fef9ec}.w.clue{background:var(--clue);border-radius:3px;padding:0 2px}
.helper{display:none;background:#eef4fb;border-left:4px solid var(--accent);padding:10px 13px;border-radius:0 8px 8px 0;margin-bottom:13px;font-size:.84rem;line-height:1.5}.helper.show{display:block}.helper strong{color:var(--primary)}.helper-cn{color:#6a4c93;font-weight:600}
.listen{border:1px solid var(--line);background:#fff;border-radius:8px;padding:6px 10px;color:var(--primary);font-size:.8rem;margin-bottom:12px}.opts{display:grid;gap:9px}.opt{display:flex;align-items:center;gap:10px;text-align:left;width:100%;padding:11px 13px;border:1.5px solid var(--line);background:#fff;border-radius:10px;font-size:.91rem}.opt:hover:not(:disabled){border-color:var(--primary);background:#f0f4fa}.letter{display:inline-flex;align-items:center;justify-content:center;width:25px;height:25px;border-radius:50%;border:1.5px solid currentColor;font-size:.75rem;font-weight:700;flex:none}.opt.correct{background:#e8f5ee;border-color:var(--ok);color:var(--ok)}.opt.wrong{background:#faeaea;border-color:var(--bad);color:var(--bad)}
.feedback{display:none;background:#f5f7fa;border:1px solid var(--line);border-radius:10px;margin-top:13px;padding:12px}.feedback.show{display:block}.fb-title{font-size:.76rem;color:var(--muted);font-weight:700;margin-bottom:8px;text-transform:uppercase}.fb-item{background:#fff;border:1px solid var(--line);border-radius:8px;padding:8px 10px;margin-top:6px;font-size:.82rem;line-height:1.45}.fb-item.answer{border-color:var(--ok);background:#e8f5ee}.fb-word{font-weight:700;color:var(--primary)}.fb-cn{color:#6a4c93;font-weight:600}
.next{display:none;width:100%;padding:12px;background:var(--primary);color:#fff;border-radius:10px;font-weight:700;margin-top:12px}.next.show{display:block}
.end{text-align:center}.grade{font-family:"Playfair Display",serif;font-size:3rem;color:var(--primary);margin:7px}.end p{color:var(--muted);line-height:1.5}.review{text-align:left;margin-top:15px;background:#f5f7fa;border:1px solid var(--line);border-radius:10px;padding:12px;font-size:.84rem;line-height:1.55}.review strong{color:var(--primary)}
.status{display:none;margin-top:12px;padding:10px;border-radius:8px;font-size:.86rem}.status.ok{display:block;color:var(--ok);background:#e8f5ee}.status.err{display:block;color:var(--bad);background:#faeaea}
.actions{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-top:14px}.actions button{padding:12px;border-radius:9px;font-weight:700}.send{background:var(--ok);color:#fff}.again{background:var(--primary);color:#fff}
#quiz,#result{display:none}@media(max-width:520px){.card{padding:15px}.actions{grid-template-columns:1fr}.title{font-size:.95rem}}
</style>
</head>
<body>
<header><div class="title">4000 Essential Words Level 5 <span>Units 1-10</span></div><div class="score">Score: <span id="score">0</span> / <span id="total">0</span></div></header>
<main class="wrap">
  <section id="start" class="card">
    <h1>Difficult Words Quiz</h1>
    <p class="sub">Review the difficult words and extended vocabulary from Units 1-10. Each question uses one of your class sentences. Tap any passage word for a definition.</p>
    <label for="studentName">Student name</label>
    <input id="studentName" maxlength="60" placeholder="Your name" autocomplete="name">
    <label for="questionCount">Number of questions</label>
    <select id="questionCount"><option value="20">20 questions</option><option value="30">30 questions</option><option value="50" selected>50 questions</option><option value="all">All available words</option></select>
    <button class="primary" onclick="startQuiz()">Start quiz</button>
  </section>
  <section id="quiz">
    <div class="progress-row"><span id="progressText"></span><span id="progressPct"></span></div>
    <div class="track"><div class="fill" id="progressFill"></div></div>
    <div class="card">
      <div class="qnum" id="qnum"></div>
      <div class="helper" id="helper"></div>
      <div class="passage" id="passage"></div>
      <button class="listen" onclick="speakCurrent()">Listen to sentence</button>
      <div class="opts" id="opts"></div>
      <div class="feedback" id="feedback"></div>
      <button class="next" id="next" onclick="nextQuestion()">Next question</button>
    </div>
  </section>
  <section id="result" class="card end">
    <h1>Quiz complete!</h1>
    <div class="grade" id="finalScore"></div>
    <p id="finalPct"></p>
    <div class="review" id="review"></div>
    <div class="status" id="status"></div>
    <div class="actions"><button class="send" onclick="submitResult()">Send result again</button><button class="again" onclick="location.reload()">Try again</button></div>
  </section>
</main>
<script>
const GOOGLE_SCRIPT_URL="https://script.google.com/macros/s/AKfycbyCD78fQsk8mKNTl0fNuRJE6AT0WpS0rAQctBeaDPGoZyd_w3T6MGNJuxnPoBCz9jIb-Q/exec";
const UNIT_NAME="4000 Essential Words Level 5 Units 1-10 Difficult Words";
const WORDS=__WORDS_JSON__;
const TESTABLE_WORDS=WORDS.filter(w=>WORDS.filter(x=>x.cat===w.cat&&x.word!==w.word).length>=3);
let questions=[],index=0,score=0,wrongWords=[],lookUpWords=[],answered=false;
const byWord=Object.fromEntries(WORDS.map(w=>[w.word.toLowerCase(),w]));
function shuffle(a){for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]]}return a}
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function startQuiz(){
  const name=document.getElementById("studentName").value.trim();
  if(!name){alert("Please enter your name first.");return}
  const raw=document.getElementById("questionCount").value;
  const count=raw==="all"?TESTABLE_WORDS.length:Math.min(Number(raw),TESTABLE_WORDS.length);
  questions=shuffle([...TESTABLE_WORDS]).slice(0,count).map(answer=>{
    const choices=shuffle(WORDS.filter(w=>w.cat===answer.cat&&w.word!==answer.word)).slice(0,3);
    return {answer,options:shuffle([answer,...choices])};
  }).filter(q=>q.options.length===4);
  document.getElementById("start").style.display="none";document.getElementById("quiz").style.display="block";
  render();
}
function clickableSentence(text){
  return String(text).split(/(_{3,}|[A-Za-z][A-Za-z'-]*)/g).map(part=>{
    if(/_{3,}/.test(part))return '<span class="blank">__________</span>';
    if(/^[A-Za-z][A-Za-z'-]*$/.test(part))return `<span class="w" data-word="${escapeHtml(part)}" onclick="lookup(this)">${escapeHtml(part)}</span>`;
    return escapeHtml(part);
  }).join("");
}
function render(){
  answered=false;const q=questions[index],pct=Math.round(index/questions.length*100);
  document.getElementById("score").textContent=score;document.getElementById("total").textContent=questions.length;
  document.getElementById("progressText").textContent=`Question ${index+1} of ${questions.length}`;document.getElementById("progressPct").textContent=pct+"%";document.getElementById("progressFill").style.width=pct+"%";document.getElementById("qnum").textContent=`Question ${index+1}`;
  document.getElementById("passage").innerHTML=clickableSentence(q.answer.sentence);
  document.getElementById("helper").className="helper";document.getElementById("feedback").className="feedback";document.getElementById("next").className="next";
  document.getElementById("opts").innerHTML=q.options.map((w,i)=>`<button class="opt" data-word="${escapeHtml(w.word)}" onclick="choose(this.dataset.word,this)"><span class="letter">${"ABCD"[i]}</span>${escapeHtml(w.word)}</button>`).join("");
}
function choose(word,btn){
  if(answered)return;answered=true;const q=questions[index],correct=word===q.answer.word;if(correct)score++;else if(!wrongWords.includes(q.answer.word))wrongWords.push(q.answer.word);
  document.getElementById("score").textContent=score;
  [...document.querySelectorAll(".opt")].forEach(b=>{b.disabled=true;const value=b.textContent.slice(1);if(value===q.answer.word)b.classList.add("correct");else if(b===btn)b.classList.add("wrong")});
  const clueWords=q.answer.sentence.replace(/_{3,}/g,"").match(/[A-Za-z][A-Za-z'-]*/g)||[];const clues=new Set(clueWords.filter(w=>w.length>5).slice(0,5).map(w=>w.toLowerCase()));document.querySelectorAll(".w").forEach(w=>{if(clues.has(w.dataset.word.toLowerCase()))w.classList.add("clue")});
  document.getElementById("feedback").innerHTML='<div class="fb-title">Vocabulary feedback: all four choices</div>'+q.options.map(w=>`<div class="fb-item ${w.word===q.answer.word?"answer":""}"><div class="fb-word">${escapeHtml(w.word)} ${w.word===q.answer.word?"(answer)":""}</div><div>${escapeHtml(w.en)}</div><div class="fb-cn">${escapeHtml(w.cn)}</div></div>`).join("");
  document.getElementById("feedback").className="feedback show";const next=document.getElementById("next");next.textContent=index===questions.length-1?"See results":"Next question";next.className="next show";
}
function nextQuestion(){index++;if(index>=questions.length)finish();else render()}
function lookup(el){
  const word=el.dataset.word.toLowerCase(),helper=document.getElementById("helper");if(word.length>2&&!lookUpWords.includes(word))lookUpWords.push(word);
  helper.className="helper show";const local=byWord[word];
  if(local){helper.innerHTML=`<strong>${escapeHtml(local.word)}</strong> <em>${escapeHtml(local.pos)}</em><br>${escapeHtml(local.en)}<br><span class="helper-cn">${escapeHtml(local.cn)}</span>`;return}
  helper.innerHTML=`<strong>${escapeHtml(word)}</strong><br><em>Looking up definition...</em>`;
  fetch("https://api.dictionaryapi.dev/api/v2/entries/en/"+encodeURIComponent(word)).then(r=>{if(!r.ok)throw 0;return r.json()}).then(data=>{const m=data[0]?.meanings?.[0];helper.innerHTML=`<strong>${escapeHtml(word)}</strong> <em>${escapeHtml(m?.partOfSpeech||"")}</em><br>${escapeHtml(m?.definitions?.[0]?.definition||"Definition not found.")}`}).catch(()=>helper.innerHTML=`<strong>${escapeHtml(word)}</strong><br><em>Definition not found. Try another word.</em>`);
}
function speak(text){if(!speechSynthesis)return;speechSynthesis.cancel();const u=new SpeechSynthesisUtterance(text);u.lang="en-US";u.rate=.9;const v=speechSynthesis.getVoices();u.voice=v.find(x=>x.lang.startsWith("en")&&/female|zira|samantha/i.test(x.name))||v.find(x=>x.lang==="en-US")||null;speechSynthesis.speak(u)}
function speakCurrent(){speak(questions[index].answer.sentence.replace(/_{3,}/g,"blank"))}
function finish(){
  document.getElementById("quiz").style.display="none";document.getElementById("result").style.display="block";const pct=Math.round(score/questions.length*100);
  document.getElementById("finalScore").textContent=`${score} / ${questions.length}`;document.getElementById("finalPct").textContent=`${pct}% correct`;
  document.getElementById("review").innerHTML=`<strong>Words to review:</strong> ${escapeHtml(wrongWords.join(", ")||"none")}<br><strong>Words looked up:</strong> ${escapeHtml(lookUpWords.join(", ")||"none")}`;
  submitResult();
}
async function submitResult(){
  const status=document.getElementById("status"),name=document.getElementById("studentName").value.trim(),pct=Math.round(score/questions.length*100),timestamp=new Date().toLocaleString("zh-TW",{timeZone:"Asia/Taipei"});
  status.className="status";status.textContent="Sending result...";
  const payload={testDateTime:timestamp,studentName:name,unitName:UNIT_NAME,percentage:pct+"%",wrongWords:wrongWords.join(", ")||"none",lookUpWords:lookUpWords.join(", ")||"none",timestamp,percent:pct,score,total:questions.length,week:UNIT_NAME,passageWords:lookUpWords.join(", ")||"none"};
  try{await fetch(GOOGLE_SCRIPT_URL,{method:"POST",mode:"no-cors",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(payload)});status.className="status ok";status.textContent="Result sent successfully. Your teacher can now see it in the spreadsheet."}
  catch(e){status.className="status err";status.textContent="Result could not be sent. Check the internet connection and tap Send result again."}
}
</script>
</body>
</html>
'@

$html = $html.Replace('__WORDS_JSON__', $json)
[IO.File]::WriteAllText($outputHtml, $html, (New-Object Text.UTF8Encoding($false)))

Write-Output "Generated: $outputHtml"
Write-Output "Extracted entries: $($items.Count)"
Write-Output "Quiz-testable entries: $testableCount"
Write-Output "Generated at: $generatedAt"
