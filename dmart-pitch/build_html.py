# -*- coding: utf-8 -*-
"""Generate a viewable, self-contained HTML deck mirroring DMart_TrolleyRest_Pitch.pptx."""
# ---- palette (from Color pallette.pdf) ----
CREAM="#F2EFE9"; DARK="#3A4C40"; SAGE="#8BA892"; CLAY="#C4694A"
SAGE_PALE="#E6ECE8"; CLAY_PALE="#F4E7E1"; INK="#2B3830"; GREY="#6E7A72"; LINE="#D8DED9"; WHITE="#FFFFFF"

# ---- component helpers (defined FIRST so SLIDES can use them) ----
def needed(t="[NEEDED]"):
    w = 23 + len(t)*15
    return f'<span class="pill" style="width:{w}px">{t}</span>'

def esc(t):
    return t.replace('&','&amp;').replace('"','&quot;')

def journey(items, hi):
    out=''
    for i,(l,s) in enumerate(items):
        active = i in hi
        out+=f'<div class="jnode {"jactive" if active else ""}"><div class="jlabel">{l}</div><div class="jsub">{s}</div></div>'
        if i<len(items)-1:
            out+='<div class="jchev">▶</div>'
    return out

def card(t,sub,chosen=False,tag=None):
    cls='card ch' if chosen else 'card'
    taghtml=f'<div class="ctag">{tag}</div>' if tag else ''
    return f'<div class="{cls}"><div class="ctitle">{t}</div><div class="csub">{sub}</div>{taghtml}</div>'

def check(t):
    return f'<div class="check"><span class="tick">✓</span><span>{t}</span></div>'

def step(t,last=False):
    cls='step sl' if last else 'step'
    return f'<div class="{cls}">{t}</div>'

def arrow():
    return '<div class="ar">→</div>'

def stepcard(num,t,d):
    return f'<div class="scard"><div class="scircle">{num}</div><div class="sctitle">{t}</div><div class="scdesc">{d}</div></div>'

def arrowbig():
    return '<div class="arbig">→</div>'

def scen(t):
    return f'<div class="scen">{t}</div>'

def pill(t):
    return f'<div class="pillbox">{t}</div>'

def krow(kpi,how):
    return f'<div class="krow"><div>{kpi}</div><div class="kneed">[NEEDED]</div><div class="kneed">[NEEDED]</div><div class="khow">{how}</div></div>'

# ---------------- SLIDES ----------------
SLIDES=[]

SLIDES.append(dict(
eyebrow="THE HOOK  ·  01",
headline="Your highest-value shoppers are the ones standing in line with nowhere to sit.",
notes="Open with the tension: DMart wins on price, but the moment that sticks with a tired "
      "shopper is the billing queue. 75% of the study's surveyed shoppers named 'nowhere to "
      "rest' as a grave problem — bigger than price or stock complaints. That is a loyalty lever "
      "DMart is not yet pulling. Set up the rest of the deck as the fix.",
inner=f'''
<div class="col" style="left:101px;top:389px;width:1066px">
  <div class="label">THE MAIN POINT</div>
  <div class="body" style="font-size:34px;line-height:1.15;margin-bottom:36px">Give tired shoppers a rest on their own trolley, and the billing-queue wait becomes a reason to come back — not a reason to rush out.</div>
  <div class="label">THE SHOPPER TRUTH</div>
  <div class="body" style="font-size:30px;color:{GREY};line-height:1.15">When asked the biggest in-store problem, shoppers did not complain about price or stock — they complained about having nowhere to sit or rest.</div>
</div>
<div class="badge" style="left:1360px;top:425px">
  <div class="bignum">75%</div>
  <div class="badge-cap">of surveyed shoppers say<br>NO REST is a grave problem</div>
  <div class="badge-tag">SURVEY SIGNAL</div>
</div>'''))

SLIDES.append(dict(
eyebrow="THE PROBLEM  ·  02",
headline="The shoppers who spend the most time and money are exactly the ones left standing.",
notes="Frame the problem in DMart's language: basket size and repeat rate. The 75% and 64% are "
      "real signals from the study's survey, but flag the sample honestly — ~50 responses from "
      "college/family, not a DMart panel. The loose-grain 64% is context, not our hero; the seating "
      "gap is. Bridge: tired shoppers want to 'bill and exit,' so the wait directly threatens "
      "unplanned purchases and loyalty.",
inner=f'''
<div class="col" style="left:101px;top:389px;width:893px">
  <div class="bullet">Older, pregnant and diverse shoppers face long billing queues in the tight space around the counters — with no resting place. The study ties this to exhaustion and a poorer store experience.</div>
  <div class="bullet">The empathy map is blunt: at some point these shoppers "just want to bill and exit." Tired = shorter browse = fewer unplanned picks.</div>
  <div class="bullet">It is a marketing problem, not just a comfort one: a worn-out exit hurts repeat intent — the metric DMart lives on.</div>
</div>
<div style="position:absolute;left:1037px;top:411px;width:778px">
  <div class="barlabel">No resting / sitting — grave problem</div>
  <div class="track"><div class="fill" style="width:75%;background:{DARK}"></div><span class="barval">75%</span></div>
  <div class="barlabel" style="margin-top:36px">Loose-grain section problems</div>
  <div class="track"><div class="fill" style="width:64%;background:{SAGE}"></div><span class="barval">64%</span></div>
  <div class="note" style="margin-top:54px;width:778px">n = ~50 survey responses (college, family &amp; relatives). Representativeness of DMart's shopper base {needed()}</div>
</div>'''))

SLIDES.append(dict(
eyebrow="THE INSIGHT  ·  03",
headline="The need to rest peaks at the billing counter — so the rest belongs there, not in a corner.",
notes="This is the pivot slide. The research already did the hard work: it located the pain "
      "precisely — ~1 hour in, at billing. That tells us a lounge or benches would backfire (more "
      "crowding). The insight is to attach the rest to the asset the shopper already has: the "
      "trolley. Keep the counter footprint unchanged. Transition into how we chose this.",
inner=f'''
<div class="journey" style="top:411px">{journey([("ARRIVE","find trolley"),("BROWSE","collect items"),("PAUSE","look for a seat"),("CHECK OUT","wait in line"),("LEAVE","bill & exit")],hi={2,3})}</div>
<div class="callout" style="top:612px">
  <div class="label" style="color:{CLAY}">THE AHA</div>
  <div class="body">The study's journey map shows the resting problem surfaces only after ~1 hour — right at billing, when the shopper is nearly done. A fixed seating zone makes crowding worse. So we bring the rest to the shopper, on the trolley, at the queue.</div>
</div>
<div class="sub" style="top:874px">Implication for DMart: the fix must live in the counter footprint — zero new floor space, zero extra queue.</div>'''))

SLIDES.append(dict(
eyebrow="IDEATION  ·  04",
headline="Four problem angles, four design rules — one resting concept cleared them all.",
notes="Show the discipline: they did not jump to a solution. Four candidate problems, then a hard "
      "constraint — staff data was unavailable, so they bet on the shopper. Among shopper problems, "
      "'no seating' won. Then four design rules acted as a filter; only a trolley-mounted rest could "
      "pass all four. That is the concept we are pitching.",
inner=f'''
<div class="grid2" style="top:396px;left:101px">
  {card("NO SEATING","no rest point",chosen=True,tag="CHOSEN")}
  {card("LOOSE-GRAIN BAGS","misplaced / confusion")}
  {card("TROLLEY PROBLEMS","wheels / manoeuvre")}
  {card("HIGH SHELVES","hard to reach")}
</div>
<div class="col" style="left:1015px;top:396px;width:778px">
  <div class="label" style="color:{CLAY}">WHY CUSTOMERS, NOT STAFF</div>
  <div class="body" style="font-size:27px;margin-bottom:30px">The team could not get primary data from staff or retail chains, so it focused on the shopper — and on "no seating" as the single main problem.</div>
  <div class="label" style="color:{CLAY}">FOUR RULES THE WINNER HAD TO PASS</div>
  {check("Independency — no help, just fold &amp; rest")}
  {check("Body support — back/arm, fits different users")}
  {check("No mess — no queue / aisle obstruction")}
  {check("Durable &amp; daily — cleanable, long-lasting")}
</div>'''))

SLIDES.append(dict(
eyebrow="PROTOTYPE  ·  05",
headline="From cardboard mock-up to a brake-locked trolley seat — we solved the one flaw that could sink it.",
notes="Proof it is real, not a sketch. The team went mock-up → Blender → mechanism POC → a built "
      "prototype envelope with concrete materials and dimensions. The critical learning: sitting "
      "shifts the cart's weight forward, so they engineered a wheel-lock brake. That is the "
      "difference between a student idea and a pilot-ready product. Next: what it actually is.",
inner=f'''
<div class="col" style="left:101px;top:389px;width:864px">
  <div class="label" style="color:{CLAY}">HOW WE BUILT IT</div>
  <div class="path">
    {step("Mini mock-up")}{arrow()}{step("Blender 3D models")}{arrow()}{step("Mechanism POCs")}{arrow()}{step("Prototype envelope",last=True)}
  </div>
  <div class="label" style="color:{CLAY};margin-top:30px">MATERIALS</div>
  <div class="bullet">Cushion: coconut fibre / firm high-density PU foam. Cover: silicon sheet or TPU. Paddle: recyclable plastic.</div>
  <div class="label" style="color:{CLAY}">SIZE &amp; FINISH</div>
  <div class="bullet">Seat 1.5–2" high, 10–12" long, 8–9" wide. Rounded edges, firm, wipe-clean surface.</div>
  <div class="label" style="color:{CLAY}">WHAT WE LEARNED</div>
  <div class="bullet">Seated weight pushed the cart forward → solved with a wheel-lock brake. Materials chosen weather / heat / water-proof and easy-clean.</div>
</div>
<div class="spec" style="left:1210px;top:425px">
  <div class="dim dim-top">H 1.5–2"</div>
  <div class="seat"><div class="seatlabel">SEAT PAD</div></div>
  <div class="dim dim-left">L 10–12"</div>
  <div class="dim dim-right">W 8–9"</div>
  <div class="speccap">Prototype envelope — built for real bodies</div>
</div>'''))

SLIDES.append(dict(
eyebrow="THE PRODUCT  ·  06",
headline="A trolley that doubles as a seat — rest at the counter, then keep shopping.",
notes="Say it in one sentence and move on — the Head of Marketing does not need mechanics, just "
      "the mental model. Three steps, all self-service: fold, sit (brake holds the cart), fold away. "
      "No staff, no new floor space, no queue change. This is the product we are asking them to pilot.",
inner=f'''
<div class="onesent">
  <span class="label" style="color:{CLAY}">IN ONE SENTENCE</span>
  <span class="sent">A foldable resting seat that clips onto the DMart trolley so shoppers sit while queuing at billing — with a wheel-lock brake for safety.</span>
</div>
<div class="steprow" style="top:584px">
  {stepcard("1","FOLD OUT","Unclip the paddle; it folds flat from the trolley frame.")}
  {arrowbig()}
  {stepcard("2","SIT + BRAKE","Shopper rests; the brake locks the wheels — no drift.")}
  {arrowbig()}
  {stepcard("3","FOLD AWAY","Clip it back; roll straight to the billing counter &amp; exit.")}
</div>'''))

feats=[("FOLD & REST","No help needed — pure self-service.","Frees floor helpers; on-brand 'everyday ease'."),
       ("BODY SUPPORT","Back & arm support; options for different users.","Serves elderly & pregnant — DMart's loyal base."),
       ("WHEEL-LOCK BRAKE","Counterweight problem, solved.","Safe in tight counter space; no aisle block."),
       ("WIPE-CLEAN BUILD","Weather / heat / water-proof, durable.","Low upkeep; meets 'cleanliness is a must'."),
       ("CLEAR OWNERSHIP","Recyclable-plastic paddle, hygienic.","Sustainable and easy to maintain daily.")]
SLIDES.append(dict(
eyebrow="KEY FUNCTIONS  ·  07",
headline="Five built-in features that turn a wait into a branded DMart moment.",
notes="Tie every feature back to a DMart benefit, not a spec. Fold-and-rest reduces calls on floor "
      "staff; body support serves the older/pregnant base DMart wants to keep; the brake is the "
      "safety answer to the prototype's weight problem; wipe-clean meets the cleanliness bar the "
      "study itself set; clear ownership keeps it hygienic and sustainable. Five reasons this is a "
      "marketing asset, not a cost.",
inner='<div class="featrow" style="top:411px">'+''.join(
    f'<div class="feat"><div class="featbar"></div><div class="featnum">{i+1:02d}</div>'
    f'<div class="feattitle">{t}</div><div class="featdesc">{d}</div>'
    f'<div class="featdiv"></div><div class="featbenefit">DMart benefit</div><div class="featb">{b}</div></div>'
    for i,(t,d,b) in enumerate(feats))+'</div>'))

SLIDES.append(dict(
eyebrow="DESIGN  ·  08",
headline="Built around real bodies and three shopper scenarios — not a generic bench.",
notes="Show this is human-centred, not decoration. The journey now has a PAUSE that appears at the "
      "counter; three real scenarios from the study are covered; and the four principles are the same "
      "rules that filtered the concept. The takeaway for marketing: this fits DMart's bodies and "
      "DMart's floor — not a showroom fantasy.",
inner=f'''
<div class="col" style="left:101px;top:382px;width:1008px">
  <div class="label" style="color:{CLAY}">REIMAGINED JOURNEY</div>
  <div class="journey" style="position:relative;top:10px">{journey([("ARRIVE","trolley"),("BROWSE","collect"),("PAUSE","seat appears"),("CHECK OUT","rest & queue"),("LEAVE","fold & exit")],hi={2,3})}</div>
  <div class="label" style="color:{CLAY};margin-top:54px">DESIGN PRINCIPLES</div>
  <div class="pills">{pill("Independency")}{pill("Body support")}{pill("No mess")}{pill("Durable & daily")}</div>
</div>
<div class="col" style="left:1195px;top:382px;width:623px">
  <div class="label" style="color:{CLAY}">THREE SCENARIOS SERVED</div>
  {scen("Independent shopper — visible rest + standing support")}
  {scen("Family companion — sit without losing the group")}
  {scen("Temporary fatigue — a short, judgement-free pause")}
  <div class="mantra">Mantra from the study: “build for real bodies.”</div>
</div>'''))

SLIDES.append(dict(
eyebrow="MEASUREMENT  ·  09",
headline="We'll track footfall, basket size, repeat rate & campaign ROI — but baselines and targets need DMart's data.",
notes="Be upfront: the source deck has zero marketing KPIs, so every baseline and target is [NEEDED] "
      "and must come from DMart's own data. What we CAN promise is the method — the same "
      "survey-based primary research that surfaced the 75%/64% signals, run pilot-vs-control. Ask "
      "the marketing head to supply the baselines so we can set real targets together.",
inner=f'''
<div class="ktable" style="left:101px;top:396px;width:1253px">
  <div class="krow khead"><div>KPI</div><div>Baseline</div><div>Target</div><div>How we measure</div></div>
  {krow("Footfall (pilot stores)","Store traffic vs control")}
  {krow("Basket size","Avg ticket at billing")}
  {krow("Repeat rate","Return visits / loyalty")}
  {krow("Campaign ROI","Spend vs incremental sales")}
</div>
<div class="col" style="left:1390px;top:396px;width:425px">
  <div class="label" style="color:{CLAY}">MEASUREMENT METHOD</div>
  <div class="bullet">Reuse the study's own method: survey-based primary research (the same approach that produced the 75% / 64% signals).</div>
  <div class="bullet">Compare pilot stores vs a control store.</div>
  <div style="margin-top:30px">{needed()}<span class="note" style="display:inline;margin-left:12px">sample &amp; cadence</span></div>
</div>'''))

asks=[("PILOT SCOPE","Rush-hour billing queues in high-footfall stores; focus on older / pregnant-heavy catchment.","store count & geographies"),
      ("TIMELINE & INVESTMENT","Build + field the prototype; measure at midpoint and close.","weeks, capex & opex"),
      ("EXPECTED OUTCOME","Validate the 75% rest-pain signal in DMart's stores; test the loyalty / queue-ease hypothesis.","quantified targets")]
SLIDES.append(dict(
eyebrow="THE ASK & NEXT STEPS  ·  10",
headline="Pilot the trolley-seat in rush-hour billing zones — scope, cost & timeline to set with DMart.",
notes="Close with a small, safe ask — not a national rollout. A rush-hour billing-zone pilot in a "
      "few high-footfall stores, using the prototype we already built. We are explicit that scope, "
      "cost and timeline are [NEEDED] from DMart, because the study gave us the problem and the "
      "product, not the P&L. Offer to co-own the measurement plan from Slide 9.",
inner=f'''
<div class="askrow" style="top:411px">
  {''.join(f'<div class="ask"><div class="askhead">{t}</div><div class="askbody">{d}</div><div class="askneed">{needed()}<span class="note" style="display:inline;margin-left:10px">{n}</span></div></div>' for t,d,n in asks)}
</div>
<div class="cta">THE ASK:&nbsp; approve a low-risk pilot &nbsp;—&nbsp; we bring the prototype, DMart brings the stores and the data.</div>'''))

# ---- assemble HTML ----
def slide_div(i,s):
    return f'''
<section class="slide" data-notes="{esc(s['notes'])}">
  <div class="topbar noanim"></div>
  <div class="chip noanim"></div>
  <div class="eyebrow">{s['eyebrow']}</div>
  <div class="headline">{s['headline']}</div>
  {s['inner']}
  <div class="footer noanim"><span>DMart · Trolley-Rest Pilot Pitch</span><span>Source: Group 7 HCD study · DMart + Star Bazaar observations</span><span class="pagenum">{i+1:02d} / 10</span></div>
</section>'''

slides_html = "\n".join(slide_div(i,s) for i,s in enumerate(SLIDES))

CSS = f"""
:root{{--cream:{CREAM};--dark:{DARK};--sage:{SAGE};--clay:{CLAY};--sagep:{SAGE_PALE};--clayp:{CLAY_PALE};--ink:{INK};--grey:{GREY};--line:{LINE};--white:{WHITE}}}
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{width:100%;height:100%;background:#222;overflow:hidden;font-family:'Work Sans',sans-serif;color:var(--ink)}}
#wrap{{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:radial-gradient(circle at 50% 30%,#2c2c2c,#161616)}}
#stage{{position:relative;width:1920px;height:1080px;background:var(--cream);overflow:hidden;box-shadow:0 30px 90px rgba(0,0,0,.5)}}
.slide{{position:absolute;inset:0;opacity:0;visibility:hidden;pointer-events:none;transition:opacity .35s ease}}
.slide.active{{opacity:1;visibility:visible;pointer-events:auto}}
@keyframes rise{{from{{opacity:0;transform:translateY(26px)}}to{{opacity:1;transform:none}}}}
@keyframes pop{{from{{opacity:0;transform:scale(.82)}}to{{opacity:1;transform:none}}}}
.noanim{{opacity:1!important}}
.topbar{{position:absolute;left:0;top:0;width:1920px;height:23px;background:var(--dark)}}
.chip{{position:absolute;left:101px;top:66px;width:43px;height:43px;background:var(--clay);border-radius:11px}}
.eyebrow{{position:absolute;left:161px;top:70px;font:700 24px 'Archivo',sans-serif;letter-spacing:1.5px;color:var(--dark)}}
.headline{{position:absolute;left:101px;top:132px;width:1720px;font:800 58px 'Archivo',sans-serif;line-height:1.04;color:var(--ink)}}
.footer{{position:absolute;left:101px;top:1452px;width:1718px;display:flex;justify-content:space-between;align-items:center;font-size:13px;color:var(--grey);border-top:1px solid var(--line);padding-top:6px}}
.pagenum{{font:700 13px 'Archivo',sans-serif;color:var(--dark)}}
.label{{font:700 24px 'Archivo',sans-serif;letter-spacing:1px;color:var(--dark);margin-bottom:8px}}
.body{{font-size:31px;line-height:1.18;color:var(--ink)}}
.bullet{{font-size:31px;line-height:1.2;color:var(--ink);margin:0 0 28px 24px;position:relative}}
.bullet::before{{content:"";position:absolute;left:-24px;top:14px;width:11px;height:11px;background:var(--clay);border-radius:3px}}
.col{{position:absolute}}
.note{{font-size:22px;color:var(--grey)}}
.sub{{position:absolute;left:101px;width:1718px;font-size:26px;color:var(--grey)}}
.badge{{position:absolute;width:439px;height:439px;border-radius:50%;background:var(--clay);display:flex;flex-direction:column;align-items:center;justify-content:center;color:#fff;text-align:center}}
.bignum{{font:800 172px 'Archivo',sans-serif;line-height:1}}
.badge-cap{{font-size:26px;font-weight:700;margin-top:6px}}
.badge-tag{{font:700 20px 'Archivo',sans-serif;letter-spacing:2px;margin-top:18px;color:#f6d9cd}}
.barlabel{{font:700 25px 'Archivo',sans-serif;color:var(--ink);margin-bottom:10px}}
.track{{position:relative;width:778px;height:72px;background:var(--sagep);border:1px solid var(--line);border-radius:14px;overflow:hidden}}
.fill{{height:72px;border-radius:14px}}
.barval{{position:absolute;right:18px;top:18px;font:700 28px 'Archivo',sans-serif;color:#fff}}
.journey{{position:absolute;left:101px;width:1718px;display:flex;align-items:center;gap:49px}}
.jnode{{flex:1;height:124px;background:#fff;border:2px solid var(--line);border-radius:18px;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center}}
.jactive{{background:var(--dark);border-color:var(--dark)}}
.jlabel{{font:700 26px 'Archivo',sans-serif;color:var(--ink)}}
.jactive .jlabel{{color:#fff}}
.jsub{{font-size:19px;color:var(--grey)}}
.jactive .jsub{{color:#cdd8cf}}
.jchev{{color:var(--sage);font-size:30px}}
.callout{{position:absolute;left:101px;width:1718px;height:223px;background:var(--sagep);border:1px solid var(--line);border-radius:18px;padding:34px 44px}}
.grid2{{position:absolute;width:893px;display:grid;grid-template-columns:1fr 1fr;gap:43px 43px}}
.card{{position:relative;background:#fff;border:2px solid var(--line);border-radius:18px;padding:32px 28px;height:223px}}
.ch{{background:var(--dark);border-color:var(--dark)}}
.ctitle{{font:700 27px 'Archivo',sans-serif;color:var(--ink)}}
.ch .ctitle{{color:#fff}}
.csub{{font-size:22px;color:var(--grey);margin-top:8px}}
.ch .csub{{color:#cdd8cf}}
.ctag{{position:absolute;top:24px;right:28px;font:700 18px 'Archivo',sans-serif;color:#fff}}
.check{{display:flex;align-items:center;gap:14px;font-size:25px;color:var(--ink);margin-bottom:16px}}
.tick{{width:49px;height:49px;border-radius:50%;background:var(--sage);color:#fff;display:flex;align-items:center;justify-content:center;font-size:26px;font-weight:700}}
.path{{display:flex;align-items:center;gap:0;margin:14px 0 0}}
.step{{height:89px;background:var(--sagep);border:2px solid var(--sage);border-radius:14px;display:flex;align-items:center;justify-content:center;text-align:center;font:700 21px 'Archivo',sans-serif;color:var(--dark);padding:0 6px}}
.sl{{background:var(--dark);color:#fff;border-color:var(--dark)}}
.ar{{color:var(--sage);font-size:30px;padding:0 10px;font-weight:700}}
.spec{{position:absolute;width:504px;height:367px}}
.seat{{position:absolute;left:72px;top:58px;width:360px;height:223px;background:var(--sagep);border:2px solid var(--sage);border-radius:22px;display:flex;align-items:center;justify-content:center}}
.seatlabel{{font:700 24px 'Archivo',sans-serif;color:var(--dark)}}
.dim{{position:absolute;font:700 22px 'Archivo',sans-serif;color:var(--ink)}}
.dim-top{{left:170px;top:-8px}}
.dim-left{{right:392px;top:150px;text-align:right;width:200px}}
.dim-right{{left:452px;top:150px}}
.speccap{{position:absolute;left:0;top:300px;width:504px;text-align:center;font-size:23px;color:var(--grey);font-style:italic}}
.onesent{{position:absolute;left:101px;top:389px;width:1718px;height:144px;background:var(--sagep);border:2px solid var(--dark);border-radius:18px;padding:22px 38px;display:flex;flex-direction:column;justify-content:center}}
.sent{{font:700 31px 'Work Sans',sans-serif;color:var(--ink);line-height:1.1}}
.steprow{{position:absolute;left:101px;width:1718px;display:flex;align-items:stretch;gap:0}}
.scard{{flex:1;background:#fff;border:1px solid var(--line);border-radius:18px;padding:40px 34px;height:324px}}
.scircle{{width:122px;height:122px;border-radius:50%;background:var(--dark);color:#fff;display:flex;align-items:center;justify-content:center;font:800 60px 'Archivo',sans-serif}}
.sctitle{{font:700 30px 'Archivo',sans-serif;color:var(--dark);margin:24px 0 14px}}
.scdesc{{font-size:25px;color:var(--grey);line-height:1.18}}
.arbig{{color:var(--sage);font-size:54px;font-weight:700;display:flex;align-items:center;padding:0 26px}}
.featrow{{position:absolute;left:101px;width:1718px;display:flex;gap:29px}}
.feat{{flex:1;background:#fff;border:1px solid var(--line);border-radius:18px;height:534px;position:relative;padding:0 26px}}
.featbar{{position:absolute;top:0;left:0;width:100%;height:17px;background:var(--dark);border-radius:18px 18px 0 0}}
.featnum{{width:121px;height:121px;border-radius:50%;background:var(--sage);color:var(--dark);display:flex;align-items:center;justify-content:center;font:800 58px 'Archivo',sans-serif;margin:42px auto 22px}}
.feattitle{{font:700 27px 'Archivo',sans-serif;color:var(--ink);text-align:center;line-height:1.05}}
.featdesc{{font-size:22px;color:var(--grey);text-align:center;margin-top:18px;line-height:1.15}}
.featdiv{{height:2px;background:var(--sage);margin:26px 0 14px}}
.featbenefit{{font:700 19px 'Archivo',sans-serif;color:var(--clay);text-align:center}}
.featb{{font-size:21px;color:var(--ink);text-align:center;margin-top:8px;line-height:1.12}}
.scen{{background:var(--sagep);border:1px solid var(--line);border-radius:14px;padding:20px 24px;font-size:24px;color:var(--ink);line-height:1.15;margin-bottom:18px}}
.mantra{{font-size:23px;color:var(--grey);font-style:italic;text-align:center;margin-top:14px}}
.pills{{display:flex;gap:20px}}
.pillbox{{background:var(--dark);color:#fff;font:700 23px 'Archivo',sans-serif;padding:12px 26px;border-radius:30px}}
.ktable{{position:absolute;border:1px solid var(--line);border-radius:14px;overflow:hidden}}
.krow{{display:grid;grid-template-columns:2.1fr 1.1fr 1.1fr 1.7fr;border-bottom:1px solid var(--line)}}
.krow:last-child{{border-bottom:none}}
.krow>div{{padding:22px 22px;font-size:25px;display:flex;align-items:center;color:var(--ink)}}
.khead{{background:var(--dark)}}
.khead>div{{color:#fff;font:700 26px 'Archivo',sans-serif}}
.kneed{{background:var(--clayp);font:700 25px 'Archivo',sans-serif;color:var(--clay)}}
.khow{{color:var(--grey);font-size:23px}}
.askrow{{position:absolute;left:101px;width:1718px;display:flex;gap:72px}}
.ask{{flex:1;background:#fff;border:1px solid var(--line);border-radius:18px;height:476px;padding:0 36px;position:relative}}
.askhead{{background:var(--dark);color:#fff;font:700 28px 'Archivo',sans-serif;text-align:center;padding:16px 0;border-radius:18px 18px 0 0;margin:0 -36px 30px}}
.askbody{{font-size:26px;color:var(--ink);line-height:1.2}}
.askneed{{position:absolute;left:36px;bottom:30px}}
.cta{{position:absolute;left:101px;top:929px;width:1718px;height:79px;background:var(--dark);border-radius:14px;color:var(--cream);display:flex;align-items:center;justify-content:center;font:700 27px 'Archivo',sans-serif}}
.pill{{display:inline-flex;align-items:center;justify-content:center;background:var(--clayp);border:1.5px solid var(--clay);color:var(--clay);font:700 22px 'Archivo',sans-serif;border-radius:20px;height:49px;padding:0 6px;vertical-align:middle}}
#nav{{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);display:flex;gap:10px;z-index:20}}
#nav button{{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.3);color:#eee;font:700 14px 'Archivo';padding:8px 16px;border-radius:8px;cursor:pointer}}
#nav button:hover{{background:rgba(255,255,255,.25)}}
#dots{{position:fixed;bottom:60px;left:50%;transform:translateX(-50%);display:flex;gap:9px;z-index:20}}
.dot{{width:11px;height:11px;border-radius:50%;background:rgba(255,255,255,.3);cursor:pointer}}
.dot.on{{background:#fff}}
#notes{{position:fixed;left:0;bottom:0;width:100%;background:rgba(20,28,23,.94);color:#e8efe9;font-size:15px;line-height:1.4;padding:14px 26px 16px;border-top:2px solid var(--clay);z-index:30;display:none}}
#notes.show{{display:block}}
#notes b{{color:var(--clay)}}
#hint{{position:fixed;top:16px;right:20px;color:rgba(255,255,255,.5);font-size:12px;z-index:20}}
"""

HTML = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DMart — Trolley-Rest Pilot Pitch</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=Work+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>{CSS}</style></head>
<body>
<div id="wrap"><div id="stage">{slides_html}</div></div>
<div id="hint">← → navigate · N notes · click sides</div>
<div id="dots"></div>
<div id="nav"><button id="prev">‹ Prev</button><button id="next">Next ›</button></div>
<div id="notes"></div>
<script>
const stage=document.getElementById('stage');
const slides=[...document.querySelectorAll('.slide')];
const dots=document.getElementById('dots');
const notesEl=document.getElementById('notes');
let cur=0;
function fit(){{const s=Math.min(innerWidth/1920,innerHeight/1080);stage.style.transform='scale('+s+')';}}
function show(i){{cur=(i+slides.length)%slides.length;slides.forEach((s,k)=>s.classList.toggle('active',k===cur));
 dots.querySelectorAll('.dot').forEach((d,k)=>d.classList.toggle('on',k===cur));
 notesEl.innerHTML='<b>Speaker notes · Slide '+(cur+1)+':</b> '+slides[cur].dataset.notes;}}
slides.forEach((_,k)=>{{const d=document.createElement('div');d.className='dot';d.onclick=()=>show(k);dots.appendChild(d);}});
function prev(){{show(cur-1)}}function next(){{show(cur+1)}}
document.getElementById('prev').onclick=prev;document.getElementById('next').onclick=next;
addEventListener('keydown',e=>{{if(e.key==='ArrowRight'||e.key===' '||e.key==='PageDown'){{next();e.preventDefault();}}
 if(e.key==='ArrowLeft'||e.key==='PageUp'){{prev();e.preventDefault();}}
 if(e.key.toLowerCase()==='n'){{notesEl.classList.toggle('show');}}}});
stage.addEventListener('click',e=>{{const r=stage.getBoundingClientRect();if(e.clientX-r.left<r.width/2)prev();else next();}});
addEventListener('resize',fit);fit();show(0);
</script></body></html>"""

# ---- post-processing: cascade entrance animations so the live preview mirrors the .pptx motion ----
HTML = HTML.replace(
  "slides[cur].dataset.notes;}",
  "slides[cur].dataset.notes; animateSlide(slides[cur]);}"
)
HTML = HTML.replace(
  "function prev(){",
  "function animateSlide(sl){const els=[...sl.querySelectorAll('*')].filter(e=>!e.classList.contains('noanim'));"
  "els.forEach((e,idx)=>{e.style.animation='none';void e.offsetWidth;"
  "const isPop=e.classList.contains('badge')||e.classList.contains('bignum')||e.classList.contains('scircle');"
  "const anim=isPop?'pop':'rise';const d=Math.min(idx*0.035,1.1);"
  "e.style.animation=anim+' .6s cubic-bezier(.2,.7,.2,1) '+d+'s both';});}\n"
  "function prev(){"
)
with open("/home/user/deck.html","w") as f:
    f.write(HTML)
print("WROTE /home/user/deck.html", len(SLIDES), "slides")
