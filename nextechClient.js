/**
 * nextechClient.js — Nextech Practice+ data layer for the Cataract Surgical Tracker.
 *
 * ⚠️  RUNS SERVER-SIDE ONLY. The Partner secret and all PHI must stay on your backend.
 *     The browser app calls YOUR backend (e.g. GET /api/board); your backend calls Nextech.
 *     Never ship clientSecret or raw FHIR responses to the React app.
 *
 * Grounded in the Nextech Practice+ API reference (FHIR STU3 / 3.0.1):
 *   Base URL : https://api.pm.nextech.com/api
 *   Auth     : OAuth2 client_credentials (Azure AD)
 *   Token    : POST https://login.microsoftonline.com/nextech-api.com/oauth2/token
 *   Headers  : Authorization: Bearer <token>   +   nx-practice-id: <practiceId>   (both required)
 *   Limits   : 20 req/s per endpoint → 429 ⇒ exponential backoff
 *   Paging   : default 10, max _count=50, follow Bundle.link[rel=next]
 *   Sync     : use _lastUpdated=ge<ts> for incremental polling
 *
 * Requires Node 18+ (global fetch).
 *
 * ──────────────────────────────────────────────────────────────────────────────
 * BOARD COLUMN  →  FHIR / Nextech SOURCE
 * ──────────────────────────────────────────────────────────────────────────────
 *   Patient name / MRN     Appointment.participant→Patient.display ; Patient.identifier (usual)
 *   Surgery (66984)        Claim item.service CPT coding (system .../cpt)  + Appointment type
 *   Eyes / laterality      Cataract dx ICD-10 laterality (H25/H26/H28 .x1/.x2/.x3) on Claim/Encounter
 *   Scheduled date         Appointment.start  (status booked|arrived|fulfilled)
 *   Preop age (92136)      Claim item.service CPT 92136 → item.servicedDateTime
 *   Clearance              DocumentReference note-category = "Medical Clearance"
 *   Auth status            APP-MANAGED — no prior-auth resource exists (Coverage = insurance only)
 *   Lens chosen / ordered  APP-MANAGED (your DB) — not a standard FHIR field
 *   Billed → drop          active Claim whose item.service CPT ∈ surgery set (status ≠ cancelled)
 *   Preop VA (outcomes)    Observation (visual acuity) — CLINICAL FHIR API, not Practice+
 * ──────────────────────────────────────────────────────────────────────────────
 */

const PM_BASE = "https://api.pm.nextech.com/api";
const TOKEN_URL = "https://login.microsoftonline.com/nextech-api.com/oauth2/token";

// CPT sets we care about
const CPT_SURGERY = ["66984", "66982", "66983", "66987", "66988"]; // phaco ± complex
const CPT_PREOP_BIOMETRY = ["92136"];
// Cataract dx prefixes (ICD-10). Laterality is the 6th char: 1=right(OD) 2=left(OS) 3=bilateral(OU)
const CATARACT_DX_PREFIX = ["H25", "H26", "H28"];

/* ───────────────────────────── auth ───────────────────────────── */

function makeClient({ practiceId, clientId, clientSecret, resource }) {
  if (!practiceId || !clientId || !clientSecret) throw new Error("nextechClient: missing credentials");
  let token = null;
  let tokenExp = 0;

  async function getToken() {
    if (token && Date.now() < tokenExp - 60_000) return token; // 60s safety margin
    const body = new URLSearchParams({
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret,
      resource, // Practice+ Partner API identifier, provided by Nextech
    });
    const res = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });
    if (!res.ok) throw new Error(`Token request failed: ${res.status}`);
    const json = await res.json();
    token = json.access_token;
    tokenExp = Date.now() + (Number(json.expires_in || 3600) * 1000);
    return token;
  }

  /* ── core GET with auth, practice header, 429 backoff, pagination ── */
  async function fhirGet(pathAndQuery, { maxPages = 20 } = {}) {
    const out = [];
    let url = pathAndQuery.startsWith("http") ? pathAndQuery : `${PM_BASE}/${pathAndQuery}`;
    for (let page = 0; page < maxPages && url; page++) {
      const bundle = await withBackoff(async () => {
        const res = await fetch(url, {
          headers: {
            Authorization: `Bearer ${await getToken()}`,
            "nx-practice-id": String(practiceId),
            Accept: "application/json",
          },
        });
        if (res.status === 429) { const e = new Error("rate-limited"); e.retry = true; throw e; }
        if (!res.ok) throw new Error(`FHIR GET ${res.status} for ${url}`);
        return res.json();
      });
      if (bundle.resourceType === "Bundle") {
        for (const e of bundle.entry || []) out.push(e.resource);
        const next = (bundle.link || []).find((l) => l.relation === "next");
        url = next ? next.url : null;
      } else {
        out.push(bundle); url = null; // single resource
      }
    }
    return out;
  }

  async function withBackoff(fn, tries = 5) {
    let delay = 400;
    for (let i = 0; i < tries; i++) {
      try { return await fn(); }
      catch (e) {
        if (!e.retry || i === tries - 1) throw e;
        await new Promise((r) => setTimeout(r, delay + Math.random() * 200));
        delay *= 2;
      }
    }
  }

  /* ───────────────────────── reference data (selectors) ───────────────────────── */

  const listPractitioners = () => fhirGet(`Practitioner?_count=50`);
  const listLocations = () => fhirGet(`Location?_count=50`);

  /* ───────────────────────── scheduling ───────────────────────── */

  // Cataract surgery appointments in a date window, optionally filtered by provider/site.
  // status booked|arrived|fulfilled = a real scheduled case.
  async function getSurgeryAppointments({ from, to, practitionerIds = [], locationIds = [] }) {
    const qp = [`date=ge${from}`, `date=lt${to}`, `status=booked,arrived,fulfilled`, `_count=50`];
    if (practitionerIds.length) qp.push(`practitioner.id=${practitionerIds.join(",")}`);
    if (locationIds.length) qp.push(`location.id=${locationIds.join(",")}`);
    const appts = await fhirGet(`Appointment?${qp.join("&")}`);
    return appts.filter(isCataractSurgeryAppt).map(parseAppointment);
  }

  // Open OR capacity: free Slots in the window (status=free). Powers "openings / first opening".
  async function getOpenSlots({ from, to, locationIds = [] }) {
    const qp = [`start=ge${from}`, `start=lt${to}`, `status=free`, `_count=50`];
    if (locationIds.length) qp.push(`location.id=${locationIds.join(",")}`);
    return (await fhirGet(`Slot?${qp.join("&")}`)).map(parseSlot);
  }

  /* ───────────────────────── billing signals (Claim) ─────────────────────────
   * CONFIRM against the Claim section of the Practice+ docs: exact search params
   * (date vs created vs billablePeriod) and where the CPT lives (item.productOrService
   * or item.service). Shapes below reflect standard STU3 Claim; adjust to Nextech's.
   * ─────────────────────────────────────────────────────────────────────────── */

  // "Preop done" log: Claims carrying CPT 92136 in the window.
  async function getPreopBiometryClaims({ from, to }) {
    const claims = await fhirGet(`Claim?created=ge${from}&created=lt${to}&_count=50`);
    return claims.filter((c) => claimHasCpt(c, CPT_PREOP_BIOMETRY)).map(parseClaim);
  }

  // "Billed → drop": Claims carrying a cataract surgery CPT. Presence ⇒ remove from board.
  async function getBilledSurgeryClaims({ from, to }) {
    const claims = await fhirGet(`Claim?created=ge${from}&created=lt${to}&_count=50`);
    return claims.filter((c) => claimHasCpt(c, CPT_SURGERY)).map(parseClaim);
  }

  /* ───────────────────────── per-patient detail ───────────────────────── */

  // Medical clearance on file = a DocumentReference categorized "Medical Clearance".
  async function getMedicalClearance(patientId) {
    const docs = await fhirGet(`Patient/${patientId}/DocumentReference?_count=50`).catch(() => []);
    const cleared = docs.find((d) => extString(d, "note-category") === "Medical Clearance" && d.status === "current");
    return { cleared: !!cleared, doc: cleared || null };
  }

  // Cataract diagnosis + laterality. Prefer the clinical FHIR Condition API; falls back to
  // Claim.diagnosis here. Returns "OU" | "OD" | "OS" | null.
  async function getCataractLaterality(patientId, { from, to }) {
    const claims = await fhirGet(`Claim?patient=${patientId}&created=ge${from}&_count=50`).catch(() => []);
    const codes = claims.flatMap(claimDxCodes).filter((c) => CATARACT_DX_PREFIX.some((p) => c.startsWith(p)));
    return lateralityFromIcd10(codes);
  }

  return {
    getToken, fhirGet,
    listPractitioners, listLocations,
    getSurgeryAppointments, getOpenSlots,
    getPreopBiometryClaims, getBilledSurgeryClaims,
    getMedicalClearance, getCataractLaterality,
    buildBoard: (opts) => buildBoard({ listPractitioners, listLocations, getSurgeryAppointments,
      getPreopBiometryClaims, getBilledSurgeryClaims, getMedicalClearance, getCataractLaterality }, opts),
  };
}

/* ───────────────────────── orchestration ─────────────────────────
 * Assembles the board the React app consumes. Joins scheduled cataract
 * surgeries with preop (92136) dates and clearance, and DROPS any patient
 * whose surgery CPT has already been billed.
 * Returns rows shaped like the app's demo `PATIENTS`.
 */
async function buildBoard(api, { from, to, practitionerIds = [], locationIds = [] }) {
  const [appts, preop, billed] = await Promise.all([
    api.getSurgeryAppointments({ from, to, practitionerIds, locationIds }),
    api.getPreopBiometryClaims({ from: shiftDays(from, -120), to }), // preop precedes surgery
    api.getBilledSurgeryClaims({ from: shiftDays(from, -30), to: shiftDays(to, 30) }),
  ]);

  const billedPatients = new Set(billed.map((c) => c.patientId));
  const preopByPatient = indexBy(preop, (c) => c.patientId); // latest 92136 per patient

  const rows = [];
  for (const a of appts) {
    if (billedPatients.has(a.patientId)) continue; // already billed → off the board

    const [clearance, laterality] = await Promise.all([
      api.getMedicalClearance(a.patientId),
      api.getCataractLaterality(a.patientId, { from: shiftDays(from, -120), to }),
    ]);
    const preopClaim = preopByPatient.get(a.patientId);

    rows.push({
      id: a.patientId,
      name: a.patientName,
      mrn: a.patientMrn,
      surgeon: a.practitionerName,
      site: a.locationName,
      surgery: "66984 · Phaco + IOL",
      laterality: laterality || "OU",
      firstEye: laterality === "OS" ? "OS" : "OD",
      scheduled: a.start,
      preopDate: preopClaim ? preopClaim.serviceDate : null,
      preopDays: preopClaim ? daysAgo(preopClaim.serviceDate) : null,
      clearance: clearance.cleared ? "cleared" : "needed",
      // ── app-managed fields (your DB), not in Practice+ FHIR ──
      auth: "pending",        // TODO: source from Coverage/Claim or your workflow store
      lens: "Pending",        // TODO: lens selection lives in your app (or IOL Assistant)
      lensOrdered: false,     // TODO: app-managed
      status: "pending",
    });
  }
  return rows;
}

/* ───────────────────────── parsers / helpers ───────────────────────── */

const refId = (ref = "") => ref.split("/").pop();
const extString = (res, urlEndsWith) =>
  (res.extension || []).find((e) => (e.url || "").endsWith(urlEndsWith))?.valueString || null;

function parseAppointment(a) {
  const parts = a.participant || [];
  const get = (type) => parts.find((p) => (p.actor?.reference || "").startsWith(type))?.actor || {};
  const patient = get("Patient"), loc = get("Location"), prac = get("Practitioner");
  return {
    apptId: a.id,
    start: a.start ? new Date(a.start) : null,
    status: a.status,
    patientId: refId(patient.reference),
    patientName: patient.display || "",
    patientMrn: "", // resolve via Patient.identifier(usual) if needed
    locationName: loc.display || "",
    locationId: refId(loc.reference),
    practitionerName: prac.display || "",
    practitionerId: refId(prac.reference),
    purpose: extRefDisplay(a, "appointment-purpose"),
    type: extRefDisplay(a, "appointment-type"),
  };
}

const extRefDisplay = (res, urlEndsWith) =>
  (res.extension || []).find((e) => (e.url || "").endsWith(urlEndsWith))?.valueReference?.display || null;

// A practice configures which appointment types/purposes mean "cataract surgery".
// "Cataract" purpose is confirmed present in Practice+; tune this to your build.
function isCataractSurgeryAppt(a) {
  const purpose = (extRefDisplay(a, "appointment-purpose") || "").toLowerCase();
  const type = (extRefDisplay(a, "appointment-type") || "").toLowerCase();
  const text = `${purpose} ${type} ${(a.description || "")}`.toLowerCase();
  return text.includes("cataract") || /surg/.test(type);
}

function parseSlot(s) {
  const contained = s.contained || [];
  const loc = contained.find((c) => c.resourceType === "Location");
  const prac = contained.find((c) => c.resourceType === "Practitioner");
  return {
    start: s.start ? new Date(s.start) : null,
    end: s.end ? new Date(s.end) : null,
    status: s.status,
    locationId: loc?.id, locationName: loc?.name,
    practitionerId: prac?.id,
  };
}

function parseClaim(c) {
  return {
    claimId: c.id,
    patientId: refId(c.patient?.reference),
    created: c.created ? new Date(c.created) : null,
    serviceDate: claimServiceDate(c),
    cpts: claimCpts(c),
    dx: claimDxCodes(c),
  };
}

// CPT codes live on Claim.item[].productOrService (STU3 may use .service). Check both.
function claimCpts(c) {
  return (c.item || []).flatMap((it) => {
    const cc = it.productOrService || it.service;
    return (cc?.coding || []).map((cd) => cd.code).filter(Boolean);
  });
}
const claimHasCpt = (c, set) => claimCpts(c).some((code) => set.includes(code));

function claimDxCodes(c) {
  return (c.diagnosis || []).flatMap((d) => (d.diagnosisCodeableConcept?.coding || []).map((cd) => cd.code)).filter(Boolean);
}
function claimServiceDate(c) {
  const d = (c.item || [])[0]?.servicedDate || (c.item || [])[0]?.servicedPeriod?.start || c.created;
  return d ? new Date(d) : null;
}

// ICD-10 cataract laterality from the trailing character: 1=OD, 2=OS, 3=bilateral
// (0/9 = unspecified). Position varies by code length (H25.13 vs H25.812), so use the last char.
function lateralityFromIcd10(codes) {
  const seen = new Set();
  for (const raw of codes) {
    const code = raw.replace(".", "");
    const lat = code[code.length - 1];
    if (lat === "3") return "OU";
    if (lat === "1") seen.add("OD");
    if (lat === "2") seen.add("OS");
  }
  if (seen.has("OD") && seen.has("OS")) return "OU";
  if (seen.has("OD")) return "OD";
  if (seen.has("OS")) return "OS";
  return null;
}

/* small utils */
const dayMs = 86_400_000;
const daysAgo = (d) => (d ? Math.round((Date.now() - +new Date(d)) / dayMs) : null);
const shiftDays = (iso, n) => new Date(+new Date(iso) + n * dayMs).toISOString().slice(0, 10);
function indexBy(arr, keyFn) {
  const m = new Map();
  for (const x of arr) { const k = keyFn(x); if (!m.has(k)) m.set(k, x); } // first (latest if pre-sorted)
  return m;
}

module.exports = { makeClient, buildBoard, lateralityFromIcd10, CPT_SURGERY, CPT_PREOP_BIOMETRY };
