// ╔══════════════════════════════════════════════════════════════╗
// ║   SoulSpark Resource Guide — Complete Apps Script  v5.0     ║
// ║   Google Apps Script · Single File for Everything           ║
// ╠══════════════════════════════════════════════════════════════╣
// ║                                                              ║
// ║  STEP 1 — SETUP (run once to build your Sheet):             ║
// ║    Click Run → runSetup                                      ║
// ║    Authorize when prompted                                   ║
// ║    All tabs created with headers, colors & validation        ║
// ║                                                              ║
// ║  STEP 2 — DEPLOY (to receive submissions from the app):      ║
// ║    Fill in TRUSTED_PIN and NOTIFY_EMAIL below                ║
// ║    Deploy → New Deployment → Web App                         ║
// ║      Execute as : Me                                         ║
// ║      Who has access : Anyone                                 ║
// ║    Your Web App URL:                                         ║
// ║    https://script.google.com/macros/s/AKfycbzg5yEU77kQPmY3x1qbF  ║
// ║    Re-deploy (new version) any time you edit this file       ║
// ║                                                              ║
// ║  CAPABILITIES ONCE DEPLOYED:                                 ║
// ║    POST  — Receives submissions from the public form         ║
// ║            Regular → Submissions tabs (pending review)       ║
// ║            Trusted PIN → Main guide tabs (live instantly)    ║
// ║    GET ?action=getSubmissions                                 ║
// ║            Returns all submissions as JSON for the Inbox     ║
// ║    GET ?action=updateStatus&section=X&rowIndex=N&status=TEXT ║
// ║            Updates a row status and colour-codes it          ║
// ║                                                              ║
// ║  SAFE TO RE-RUN runSetup: existing tabs are never deleted    ║
// ╚══════════════════════════════════════════════════════════════╝


// ══════════════════════════════════════════════════════════════
//  CONFIGURATION  ← edit these before deploying
// ══════════════════════════════════════════════════════════════

// Share this PIN with trusted collaborators. Their submissions
// skip the review queue and go straight into the main guide tabs.
// Set to "" to disable the trusted PIN feature entirely.
const TRUSTED_PIN = "";          // e.g.  "SPARK2025"

// Email address for new-submission notifications.
// Set to "" to disable email notifications entirely.
const NOTIFY_EMAIL = "";         // e.g.  "team@soulspark.com"

// Your deployed Web App URL — paste this into the SoulSpark app and Submission Form
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzg5yEU77kQPmY3x1qbFORpwPqco9C3Oghj-vE7DN2NUbY7WTz-Vz8IPkW5Jg9Sm1o/exec";


// ══════════════════════════════════════════════════════════════
//  TAB NAMES
// ══════════════════════════════════════════════════════════════

// Where regular (pending review) submissions land.
const SUBMISSION_TABS = {
  people:     "👥 Submissions — People",
  networking: "🤝 Submissions — Networking",
  shows:      "🎪 Submissions — Shows",
  expos:      "🎡 Submissions — Expos",
  modalities: "🔮 Submissions — Modalities",
  stores:     "🏪 Submissions — Stores",
  links:      "🔗 Submissions — Links",
  ideas:      "💡 Submissions — Ideas",
  misc:       "📚 Submissions — Vault",
  events:     "📅 Submissions — Events",
};

// Where trusted-PIN submissions land (straight into the live guide).
const MAIN_TABS = {
  people:     "People & Contacts",
  networking: "Networking",
  shows:      "Shows & Events",
  expos:      "Expos & Markets",
  modalities: "Modalities",
  stores:     "Stores & Vendors",
  links:      "Useful Links",
  ideas:      "Ideas & Experiments",
  misc:       "Resource Vault",
  events:     "Upcoming Events",
};

// Tab accent colours.
const TAB_COLORS = {
  people:     "#c47a1a",
  networking: "#2d5a27",
  shows:      "#1f6670",
  expos:      "#1f6670",
  modalities: "#5a2db8",
  stores:     "#a02020",
  links:      "#4a4a6a",
  ideas:      "#1a6090",
  misc:       "#8b3a10",
  events:     "#6b2fa0",
};

// Status dropdown options (used in data validation).
const STATUS_OPTIONS = [
  "⏳ Pending Review",
  "✅ Approved — Added",
  "❌ Rejected",
  "🔁 Needs More Info",
];


// ══════════════════════════════════════════════════════════════
//  COLUMN HEADERS  (must stay in sync with FIELD_MAP below)
// ══════════════════════════════════════════════════════════════

const HEADERS = {

  people: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    // ── Identity ──────────────────────────────────────────────
    "Full Name", "Pronouns", "Role / Title", "Org / Business", "Location",
    // ── Contact ───────────────────────────────────────────────
    "Phone", "Email",
    // ── Service ───────────────────────────────────────────────
    "About / Bio", "Modalities", "Service Area",
    "Serves Local", "Serves Regional", "Serves National",
    "Serves International", "Serves Online", "Serves In-Person",
    // ── Metaphysical ──────────────────────────────────────────
    "MBTI", "HD Type", "Enneagram",
    "Sun Sign", "Moon Sign", "Rising Sign",
    // ── Social ────────────────────────────────────────────────
    // ── Social & Web ──────────────────────────────────────────
    "Website", "Instagram", "TikTok", "Facebook", "LinkedIn",
    "YouTube", "Threads", "X / Twitter", "Pinterest",
    "Substack", "Patreon", "Etsy", "Shopify", "Telegram", "Signal",
    "Other Social 1 Label", "Other Social 1",
    "Other Social 2 Label", "Other Social 2",
    "Other Social 3 Label", "Other Social 3",
    "Other Social 4 Label", "Other Social 4",
    "Other Social 5 Label", "Other Social 5",
    // ── Experience & Seeking ──────────────────────────────────
    "Experience Level",
    "Seeking — Mentor", "Seeking — Students", "Seeking — Collaborators",
    "Seeking — Resources", "Seeking — Referrals",
    // ── Ratings & Notes ───────────────────────────────────────
    "Recommend?", "Rating", "Rec. Comment", "Notes",
  ],

  networking: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Org / Group Name", "Focus / Industry", "Location",
    "Website", "Contact Person", "Email / Phone",
    "Notes",
  ],

  shows: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Event Name", "Location", "Website", "Organizer Contact",
    "Type — Psychic", "Type — Wellness", "Type — Craft",
    "Type — Spiritual", "Type — Holistic",
    "Date(s)", "Booth Fee", "Setup Size", "Electricity",
    "Worth Again?", "Event Notes",
  ],

  expos: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Expo / Market Name", "Location", "Website", "Organizer Contact",
    "Date(s)", "Booth Fee", "Setup Size",
    "Electricity Included", "Electricity Extra",
    "Worth Again?", "Expo Notes",
  ],

  modalities: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Modality Name",
    "Category — Energy", "Category — Divination", "Category — Somatic",
    "Category — Coaching", "Category — Creative",
    "Description", "Ideal Client", "Tools Used",
    "Certification", "Resources / Suppliers", "Recommend?",
  ],

  stores: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Store Name", "Location", "Website / Instagram", "Buyer / Manager",
    "Type — Metaphysical", "Type — Boutique", "Type — Wellness",
    "Type — Art", "Type — Supplier",
    "Opp — Consignment", "Opp — Wholesale", "Opp — Pop-Up", "Opp — Permanent",
    "Contact Info", "Submission Requirements", "Follow-Up Date", "Additional Notes",
  ],

  links: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Resource Name", "Category", "URL", "Description", "Recommend?",
  ],

  ideas: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Idea Name", "Intended Audience", "Idea Description",
    "Type — Offer", "Type — Digital", "Type — Content",
    "Type — Collab", "Type — Event", "Type — System",
    "Next Steps", "Personal Notes",
  ],

  misc: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Title / Resource Name", "Author / Creator",
    "Cat — Books", "Cat — Course", "Cat — YouTube", "Cat — Media",
    "Cat — Intuitive", "Cat — App", "Cat — Website", "Cat — Idea",
    "Other Category",
    "Description / Summary", "Link / Access Info", "Why Save This?",
    "Image (Drive Link)",
  ],

  events: [
    "Submitted At", "Submitter Name", "Submitter Email", "Status",
    "Event Name", "Event Type", "Start Date", "End Date",
    "Apply / Vend Deadline", "Venue", "City", "State", "Zip",
    "Address", "Organizer", "Website", "Apply to Vend URL",
    "Cost to Attend", "Vend / Booth Fee", "Setup Size", "Electricity",
    "Modalities Featured", "Description", "Associated With",
    "Flyer (Drive Link)",
  ],

};


// ══════════════════════════════════════════════════════════════
//  FIELD MAP  — form field name  →  column header string
//  Keys must match the `name` attributes sent by the form.
// ══════════════════════════════════════════════════════════════

const FIELD_MAP = {

  people: {
    name:                  "Full Name",
    pronouns:              "Pronouns",
    role_title:            "Role / Title",
    org:                   "Org / Business",
    location:              "Location",
    phone:                 "Phone",
    email:                 "Email",
    bio:                   "About / Bio",
    modalities:            "Modalities",
    service_area:          "Service Area",
    serves_local:          "Serves Local",
    serves_regional:       "Serves Regional",
    serves_national:       "Serves National",
    serves_international:  "Serves International",
    serves_online:         "Serves Online",
    serves_inperson:       "Serves In-Person",
    mbti:                  "MBTI",
    hd_type:               "HD Type",
    enneagram:             "Enneagram",
    sun_sign:              "Sun Sign",
    moon_sign:             "Moon Sign",
    rising_sign:           "Rising Sign",
    website:               "Website",
    instagram:             "Instagram",
    tiktok:                "TikTok",
    facebook:              "Facebook",
    linkedin:              "LinkedIn",
    youtube:               "YouTube",
    threads:               "Threads",
    twitter:               "X / Twitter",
    pinterest:             "Pinterest",
    substack:              "Substack",
    patreon:               "Patreon",
    etsy:                  "Etsy",
    shopify:               "Shopify",
    telegram:              "Telegram",
    signal:                "Signal",
    other_social_1_label:  "Other Social 1 Label",
    other_social_1:        "Other Social 1",
    other_social_2_label:  "Other Social 2 Label",
    other_social_2:        "Other Social 2",
    other_social_3_label:  "Other Social 3 Label",
    other_social_3:        "Other Social 3",
    other_social_4_label:  "Other Social 4 Label",
    other_social_4:        "Other Social 4",
    other_social_5_label:  "Other Social 5 Label",
    other_social_5:        "Other Social 5",
    exp_level:             "Experience Level",
    seeking_mentor:        "Seeking — Mentor",
    seeking_student:       "Seeking — Students",
    seeking_collab:        "Seeking — Collaborators",
    seeking_resources:     "Seeking — Resources",
    seeking_referrals:     "Seeking — Referrals",
    recommend:             "Recommend?",
    rating:                "Rating",
    rec_comment:           "Rec. Comment",
    notes:                 "Notes",
    submitter_name:        "Submitter Name",
    submitter_email:       "Submitter Email",
  },

  networking: {
    org_name:        "Org / Group Name",
    focus:           "Focus / Industry",
    location:        "Location",
    website:         "Website",
    contact:         "Contact Person",
    email_phone:     "Email / Phone",
    notes:           "Notes",
    submitter_name:  "Submitter Name",
    submitter_email: "Submitter Email",
  },

  shows: {
    event_name:      "Event Name",
    location:        "Location",
    website:         "Website",
    organizer:       "Organizer Contact",
    type_psychic:    "Type — Psychic",
    type_wellness:   "Type — Wellness",
    type_craft:      "Type — Craft",
    type_spiritual:  "Type — Spiritual",
    type_holistic:   "Type — Holistic",
    dates:           "Date(s)",
    booth_fee:       "Booth Fee",
    setup_size:      "Setup Size",
    electricity:     "Electricity",
    worth:           "Worth Again?",
    event_notes:     "Event Notes",
    submitter_name:  "Submitter Name",
    submitter_email: "Submitter Email",
  },

  expos: {
    expo_name:       "Expo / Market Name",
    location:        "Location",
    website:         "Website",
    organizer:       "Organizer Contact",
    dates:           "Date(s)",
    booth_fee:       "Booth Fee",
    setup_size:      "Setup Size",
    elec_included:   "Electricity Included",
    elec_extra:      "Electricity Extra",
    worth:           "Worth Again?",
    expo_notes:      "Expo Notes",
    submitter_name:  "Submitter Name",
    submitter_email: "Submitter Email",
  },

  modalities: {
    modality_name:        "Modality Name",
    cat_energy:           "Category — Energy",
    cat_divination:       "Category — Divination",
    cat_somatic:          "Category — Somatic",
    cat_coaching:         "Category — Coaching",
    cat_creative:         "Category — Creative",
    description:          "Description",
    ideal_client:         "Ideal Client",
    tools_used:           "Tools Used",
    certification:        "Certification",
    resources_suppliers:  "Resources / Suppliers",
    recommend:            "Recommend?",
    submitter_name:       "Submitter Name",
    submitter_email:      "Submitter Email",
  },

  stores: {
    store_name:        "Store Name",
    location:          "Location",
    website:           "Website / Instagram",
    buyer_name:        "Buyer / Manager",
    type_metaphysical: "Type — Metaphysical",
    type_boutique:     "Type — Boutique",
    type_wellness:     "Type — Wellness",
    type_art:          "Type — Art",
    type_supplier:     "Type — Supplier",
    opp_consignment:   "Opp — Consignment",
    opp_wholesale:     "Opp — Wholesale",
    opp_popup:         "Opp — Pop-Up",
    opp_permanent:     "Opp — Permanent",
    contact_info:      "Contact Info",
    submission_req:    "Submission Requirements",
    followup_date:     "Follow-Up Date",
    additional_info:   "Additional Notes",
    submitter_name:    "Submitter Name",
    submitter_email:   "Submitter Email",
  },

  links: {
    link_name:       "Resource Name",
    link_category:   "Category",
    link_url:        "URL",
    link_desc:       "Description",
    recommend:       "Recommend?",
    submitter_name:  "Submitter Name",
    submitter_email: "Submitter Email",
  },

  ideas: {
    idea_name:        "Idea Name",
    audience:         "Intended Audience",
    idea_description: "Idea Description",
    itype_offer:      "Type — Offer",
    itype_digital:    "Type — Digital",
    itype_content:    "Type — Content",
    itype_collab:     "Type — Collab",
    itype_event:      "Type — Event",
    itype_system:     "Type — System",
    summary:          "Next Steps",
    personal_notes:   "Personal Notes",
    submitter_name:   "Submitter Name",
    submitter_email:  "Submitter Email",
  },

  misc: {
    title:             "Title / Resource Name",
    creator:           "Author / Creator",
    cat_books:         "Cat — Books",
    cat_course:        "Cat — Course",
    cat_youtube:       "Cat — YouTube",
    cat_media:         "Cat — Media",
    cat_intuitive:     "Cat — Intuitive",
    cat_software:      "Cat — App",
    cat_website:       "Cat — Website",
    cat_idea:          "Cat — Idea",
    cat_other_text:    "Other Category",
    description:       "Description / Summary",
    access_info:       "Link / Access Info",
    why_stored:        "Why Save This?",
    _vault_image_link: "Image (Drive Link)",
    submitter_name:    "Submitter Name",
    submitter_email:   "Submitter Email",
  },

  events: {
    event_name:       "Event Name",
    event_type:       "Event Type",
    start_date:       "Start Date",
    end_date:         "End Date",
    apply_deadline:   "Apply / Vend Deadline",
    venue:            "Venue",
    city:             "City",
    state:            "State",
    zip:              "Zip",
    address:          "Address",
    organizer:        "Organizer",
    website:          "Website",
    apply_url:        "Apply to Vend URL",
    cost_attend:      "Cost to Attend",
    booth_fee:        "Vend / Booth Fee",
    setup_size:       "Setup Size",
    electricity:      "Electricity",
    modalities:       "Modalities Featured",
    description:      "Description",
    associated_with:  "Associated With",
    _flyer_link:      "Flyer (Drive Link)",
    submitter_name:   "Submitter Name",
    submitter_email:  "Submitter Email",
  },

};


// ══════════════════════════════════════════════════════════════
//  STEP 1 — runSetup
//  Run this once to build all tabs. Safe to re-run — existing
//  tabs are skipped, never deleted or overwritten.
// ══════════════════════════════════════════════════════════════

function runSetup() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  log.push("╔══════════════════════════════════════╗");
  log.push("║   SoulSpark Sheet Setup Starting…   ║");
  log.push("╚══════════════════════════════════════╝");
  log.push("");

  // ── 1. Submission tabs ───────────────────────────────────────
  log.push("── Submission Tabs ──────────────────────");
  Object.entries(SUBMISSION_TABS).forEach(([sec, name]) => {
    const r = buildTab(ss, name, sec, true);
    log.push((r.created ? "  ✓ Created  " : "  → Exists  ") + name);
  });

  // ── 2. Main guide tabs ───────────────────────────────────────
  log.push("");
  log.push("── Main Guide Tabs ──────────────────────");
  Object.entries(MAIN_TABS).forEach(([sec, name]) => {
    const r = buildTab(ss, name, sec, false);
    log.push((r.created ? "  ✓ Created  " : "  → Exists  ") + name);
  });

  // ── 3. Dashboard tab ─────────────────────────────────────────
  log.push("");
  log.push("── Dashboard ────────────────────────────");
  const dash = buildDashboard(ss);
  log.push(dash ? "  ✓ Created  📊 Dashboard" : "  → Exists  📊 Dashboard");

  // ── 4. Move Dashboard to first position ──────────────────────
  try {
    const dashSheet = ss.getSheetByName("📊 Dashboard");
    if (dashSheet) { ss.setActiveSheet(dashSheet); ss.moveActiveSheet(1); }
  } catch(e) {}

  // ── 5. Summary ───────────────────────────────────────────────
  const totalTabs = Object.keys(SUBMISSION_TABS).length + Object.keys(MAIN_TABS).length;
  log.push("");
  log.push("══════════════════════════════════════════");
  log.push("  Setup complete! ✦");
  log.push("  Next: deploy this script as a Web App,");
  log.push("  then paste the URL into the SoulSpark");
  log.push("  Your Web App URL: https://script.google.com/macros/s/AKfycbzg5yEU77kQPmY3x1qbFORpwPqco9C3Oghj-vE7DN2NUbY7WTz-Vz8IPkW5Jg9Sm1o/exec");
  log.push("  Paste it into the SoulSpark app under 'Connect Google Sheets'.");
  log.push("══════════════════════════════════════════");

  Logger.log(log.join("\n"));
  SpreadsheetApp.getUi().alert(
    "✦ SoulSpark Setup Complete!",
    totalTabs + " tabs are ready.\n\nYour Web App URL:\nhttps://script.google.com/macros/s/AKfycbzg5yEU77kQPmY3x1qbFORpwPqco9C3Oghj-vE7DN2NUbY7WTz-Vz8IPkW5Jg9Sm1o/exec\n\nNext step (if not already connected): Deploy this script as a Web App (Deploy → New Deployment → Web App, Execute as: Me, Who has access: Anyone), then copy the Web App URL and paste it into the SoulSpark app under 'Connect Google Sheets'.\n\nSee the 📊 Dashboard tab for full instructions.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════
//  STEP 2 — doPost  (Web App endpoint — receives submissions)
// ══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const raw = JSON.parse(e.postData.contents);
    const sec = (raw._section || "").toLowerCase().trim();

    // Validate section
    if (!HEADERS[sec]) {
      return respond({ status: "error", message: "Unknown section: " + sec });
    }

    // Trusted PIN check
    const submittedPin = String(raw._trusted_pin || "").trim().toUpperCase();
    const configPin    = String(TRUSTED_PIN      || "").trim().toUpperCase();
    const isTrusted    = !!(configPin && submittedPin === configPin);

    const tabName     = isTrusted ? (MAIN_TABS[sec] || sec) : (SUBMISSION_TABS[sec] || sec);
    const statusValue = isTrusted ? "✅ Auto-Approved (Trusted PIN)" : "⏳ Pending Review";

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, tabName, sec, isTrusted);

    // Vault image: encode to Drive and inject link
    if (sec === "misc" && raw.vault_image_b64) {
      try {
        raw._vault_image_link = saveImageToDrive(
          raw.vault_image_b64,
          raw.vault_image_mime     || "image/jpeg",
          raw.vault_image_filename || ("vault_" + Date.now() + ".jpg"),
          raw.title || "Untitled"
        );
      } catch (imgErr) {
        raw._vault_image_link = "⚠ Image save failed: " + imgErr.message;
      }
      delete raw.vault_image_b64;
      delete raw.vault_image_mime;
      delete raw.vault_image_filename;
    }

    // Events flyer: encode to Drive and inject link
    if (sec === "events" && raw.flyer_b64) {
      try {
        raw._flyer_link = saveImageToDrive(
          raw.flyer_b64,
          raw.flyer_mime     || "image/jpeg",
          raw.flyer_filename || ("flyer_" + Date.now() + ".jpg"),
          raw.event_name || "Event"
        );
      } catch (imgErr) {
        raw._flyer_link = "⚠ Flyer save failed: " + imgErr.message;
      }
      delete raw.flyer_b64;
      delete raw.flyer_mime;
      delete raw.flyer_filename;
    }

    // Strip internal fields before building the row
    delete raw._trusted_pin;
    delete raw._section;

    // Build row array aligned to this section's headers
    const map    = FIELD_MAP[sec];
    const hdrs   = HEADERS[sec];
    const rowArr = hdrs.map(h => {
      if (h === "Submitted At") return new Date().toLocaleString("en-US");
      if (h === "Status")       return statusValue;
      const field = Object.keys(map).find(k => map[k] === h);
      if (!field) return "";
      const val = raw[field];
      if (val === undefined || val === null) return "";
      if (val === true  || val === "true")   return "✓";
      if (val === false || val === "false")  return "";
      return String(val);
    });

    sheet.appendRow(rowArr);
    notifyTeam(sec, raw, isTrusted);

    return respond({ status: "success", section: sec, trusted: isTrusted });

  } catch (err) {
    return respond({ status: "error", message: err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  doGet  —  Inbox API endpoints
// ══════════════════════════════════════════════════════════════
//
//  ?action=getSubmissions
//      Returns every row from every Submissions tab as a flat
//      JSON array, sorted pending-first. Powers the Inbox tab.
//
//  ?action=updateStatus&section=X&rowIndex=N&status=TEXT
//      Writes a new status string and colour-codes the row.
//
//  (no params)  →  health-check response
// ══════════════════════════════════════════════════════════════

function doGet(e) {
  const p        = (e && e.parameter) ? e.parameter : {};
  const action   = (p.action || "").trim();
  const callback = (p.callback || "").trim();   // JSONP support

  if (action === "getSubmissions") return getSubmissions(callback);

  if (action === "getDirectory") return getDirectory(callback);

  // ?action=getSection&section=people  → returns main-tab rows for any section
  if (action === "getSection") {
    const sec = (p.section || "people").trim();
    return getSection(sec, callback);
  }

  // ?action=getAllSections → returns all main-tab rows merged (for mandala/binder)
  if (action === "getAllSections") return getAllSections(callback);

  if (action === "updateStatus") {
    return updateStatus(
      (p.section  || "").trim(),
      parseInt(p.rowIndex || "0", 10),
      (p.status   || "").trim()
    );
  }

  return respond({
    status:  "ok",
    message: "SoulSpark Endpoint is live ✦",
    time:    new Date().toLocaleString("en-US"),
  });
}


// ══════════════════════════════════════════════════════════════
//  getSubmissions
// ══════════════════════════════════════════════════════════════

function getSubmissions(callback) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const all = [];

    Object.entries(SUBMISSION_TABS).forEach(([sec, tabName]) => {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet) return;
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      const data    = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
      const headers = data[0];

      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        if (row.every(v => String(v).trim() === "")) continue;

        const obj = { _section: sec, _rowIndex: r + 1 };
        headers.forEach((h, c) => {
          obj[headerToSnake(h)] = row[c] !== undefined ? String(row[c]) : "";
        });
        obj.submitted_at    = obj.submitted_at    || "";
        obj.submitter_name  = obj.submitter_name  || "";
        obj.submitter_email = obj.submitter_email || "";
        obj.status          = obj.status          || "⏳ Pending Review";
        all.push(obj);
      }
    });

    all.sort((a, b) => {
      return (a.status.startsWith("⏳") ? 0 : 1) - (b.status.startsWith("⏳") ? 0 : 1);
    });

    return respond({ status: "ok", submissions: all, count: all.length });

  } catch (err) {
    return respond({ status: "error", message: err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  updateStatus
// ══════════════════════════════════════════════════════════════
//  getDirectory  —  returns approved People entries for public directory
// ══════════════════════════════════════════════════════════════

function getDirectory(callback) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MAIN_TABS["people"]);
    if (!sheet) return respond({ status: "ok", people: [], count: 0 });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return respond({ status: "ok", people: [], count: 0 });

    const data    = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const headers = data[0];

    // Fields to expose publicly (exclude private/internal fields)
    const PUBLIC_FIELDS = new Set([
      "full_name","role_title","pronouns","org_business","location",
      "about_bio","modalities","service_area",
      "serves_local","serves_regional","serves_national","serves_international","serves_online","serves_in_person",
      "mbti","hd_type","enneagram","sun_sign","moon_sign","rising_sign",
      "instagram","website","tiktok","facebook","linkedin",
      "experience_level","recommend","rating",
      "photo_b64","photo_mime","logo_b64","logo_mime",
      "specialty_tags","service_types","role_tags","ideal_client","certifications","classes",
      "descriptives","testimonial","collab_notes","endorsements"
    ]);

    const people = [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (row.every(v => String(v).trim() === "")) continue;
      const obj = {};
      headers.forEach((h, c) => {
        const key = headerToSnake(h);
        if (PUBLIC_FIELDS.has(key)) {
          obj[key] = row[c] !== undefined ? String(row[c]) : "";
        }
      });
      if (obj.full_name) people.push(obj);
    }

    return respond({ status: "ok", people: people, count: people.length }, callback);

  } catch (err) {
    return respond({ status: "error", message: err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════


// ══════════════════════════════════════════════════════════════
//  getSection — returns rows from any main-tab section as JSON
// ══════════════════════════════════════════════════════════════
function getSection(sec, callback) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const tabMap  = Object.assign({}, MAIN_TABS);
    const tabName = tabMap[sec] || tabMap["people"];
    const sheet   = ss.getSheetByName(tabName);
    if (!sheet) return respond({ status: "ok", rows: [], count: 0, section: sec });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return respond({ status: "ok", rows: [], count: 0, section: sec });

    const data    = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const headers = data[0];
    const rows    = [];

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (row.every(v => String(v).trim() === "")) continue;
      const obj = { _section: sec };
      headers.forEach((h, c) => {
        obj[headerToSnake(h)] = row[c] !== undefined ? String(row[c]) : "";
      });
      rows.push(obj);
    }

    return respond({ status: "ok", rows: rows, count: rows.length, section: sec });
  } catch(err) {
    return respond({ status: "error", message: err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════
//  getAllSections — returns ALL main-tab rows merged for mandala/binder
// ══════════════════════════════════════════════════════════════
function getAllSections(callback) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const all = [];

    Object.entries(MAIN_TABS).forEach(([sec, tabName]) => {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet) return;
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      const data    = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
      const headers = data[0];

      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        if (row.every(v => String(v).trim() === "")) continue;
        const obj = { _section: sec };
        headers.forEach((h, c) => {
          obj[headerToSnake(h)] = row[c] !== undefined ? String(row[c]) : "";
        });
        if (obj.full_name || obj.event_name || obj.store_name || obj.modality_name || obj.org_group_name || obj.idea_name || obj.title_resource_name || obj.resource_name) {
          all.push(obj);
        }
      }
    });

    return respond({ status: "ok", rows: all, count: all.length });
  } catch(err) {
    return respond({ status: "error", message: err.toString() });
  }
}

function updateStatus(sec, rowIndex, newStatus) {
  try {
    if (!sec || rowIndex < 2 || !newStatus) {
      return respond({ status: "error", message: "Missing or invalid parameter. Need: section, rowIndex (≥2), status." });
    }

    const tabName = SUBMISSION_TABS[sec];
    if (!tabName) return respond({ status: "error", message: "Unknown section: " + sec });

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return respond({ status: "error", message: "Tab not found: " + tabName });

    const firstRow  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = firstRow.indexOf("Status") + 1;
    if (statusCol < 1) return respond({ status: "error", message: "Status column not found in: " + tabName });

    sheet.getRange(rowIndex, statusCol).setValue(newStatus);

    const rowRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
    if      (newStatus.startsWith("✅")) rowRange.setBackground("#e8f0e6");
    else if (newStatus.startsWith("❌")) rowRange.setBackground("#fce8e8");
    else if (newStatus.startsWith("🔁")) rowRange.setBackground("#ddeeff");
    else if (newStatus.startsWith("⏳")) rowRange.setBackground("#fdf0dc");
    else                                  rowRange.setBackground(null);

    return respond({ status: "ok", updated: rowIndex, newStatus: newStatus });

  } catch (err) {
    return respond({ status: "error", message: err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  buildTab  —  creates a sheet with full formatting
//  Used by both runSetup and getOrCreateSheet (on-demand).
// ══════════════════════════════════════════════════════════════

function buildTab(ss, tabName, sec, isSubmissions) {
  if (ss.getSheetByName(tabName)) return { created: false };

  const sheet   = ss.insertSheet(tabName);
  const headers = HEADERS[sec] || [];

  // Header row
  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setValues([headers]);
  hdrRange.setFontWeight("bold");
  hdrRange.setFontSize(10);
  hdrRange.setFontFamily("Arial");
  hdrRange.setBackground(isSubmissions ? "#f0eee1" : "#e8f0e6");
  hdrRange.setFontColor("#1a1410");
  hdrRange.setBorder(null, null, true, null, null, null, "#d6cfc0", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);

  // Column widths
  sheet.setColumnWidth(1, 150);  // Submitted At
  sheet.setColumnWidth(2, 130);  // Submitter Name
  sheet.setColumnWidth(3, 160);  // Submitter Email
  sheet.setColumnWidth(4, 190);  // Status
  const maxCol = Math.min(headers.length, 25);
  for (let c = 5; c <= maxCol; c++) sheet.setColumnWidth(c, 140);

  // Status data validation dropdown
  const statusCol = headers.indexOf("Status") + 1;
  if (statusCol > 0 && sheet.getMaxRows() > 1) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_OPTIONS, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, statusCol, 999, 1).setDataValidation(rule);
  }

  // Conditional formatting (colour-coded by status emoji)
  if (statusCol > 0) {
    const cfRange = sheet.getRange(2, 1, 999, headers.length);
    const rules   = [
      ["⏳", "#fdf0dc"],
      ["✅", "#e8f0e6"],
      ["❌", "#fce8e8"],
      ["🔁", "#ddeeff"],
    ].map(([emoji, bg]) =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=LEFT(INDIRECT("R"&ROW()&"C${statusCol}",FALSE),2)="${emoji}"`
        )
        .setBackground(bg)
        .setRanges([cfRange])
        .build()
    );
    sheet.setConditionalFormatRules(rules);
  }

  // Tab color
  if (TAB_COLORS[sec]) sheet.setTabColor(TAB_COLORS[sec]);

  return { created: true };
}


// ══════════════════════════════════════════════════════════════
//  getOrCreateSheet  —  called by doPost when receiving a live
//  submission. Builds the tab on-demand if it doesn't exist yet.
// ══════════════════════════════════════════════════════════════

function getOrCreateSheet(ss, tabName, sec, isTrusted) {
  const existing = ss.getSheetByName(tabName);
  if (existing) return existing;
  buildTab(ss, tabName, sec, !isTrusted);
  return ss.getSheetByName(tabName);
}


// ══════════════════════════════════════════════════════════════
//  buildDashboard  —  creates the 📊 Dashboard / README tab
// ══════════════════════════════════════════════════════════════

function buildDashboard(ss) {
  if (ss.getSheetByName("📊 Dashboard")) return false;

  const sheet = ss.insertSheet("📊 Dashboard");
  sheet.setTabColor("#6b2fa0");

  sheet.getRange("A1").setValue("✦ SoulSpark Resource Guide");
  sheet.getRange("A1").setFontSize(20).setFontWeight("bold").setFontColor("#6b2fa0").setFontFamily("Georgia");
  sheet.getRange("A2").setValue("Google Sheets Backend — Connected via Apps Script");
  sheet.getRange("A2").setFontSize(11).setFontColor("#6b6258");
  sheet.getRange("A3").setValue("Last setup run: " + new Date().toLocaleString("en-US"));
  sheet.getRange("A3").setFontSize(10).setFontColor("#a09688");

  const sections = [
    ["Tab Name", "Purpose", "Who Writes Here"],
    ["👥 Submissions — People",     "People submissions awaiting review",     "Public form → App Inbox"],
    ["🤝 Submissions — Networking", "Networking submissions awaiting review", "Public form → App Inbox"],
    ["🎪 Submissions — Shows",      "Show submissions awaiting review",       "Public form → App Inbox"],
    ["🎡 Submissions — Expos",      "Expo submissions awaiting review",       "Public form → App Inbox"],
    ["🔮 Submissions — Modalities", "Modality submissions awaiting review",   "Public form → App Inbox"],
    ["🏪 Submissions — Stores",     "Store submissions awaiting review",      "Public form → App Inbox"],
    ["🔗 Submissions — Links",      "Link submissions awaiting review",       "Public form → App Inbox"],
    ["💡 Submissions — Ideas",      "Idea submissions awaiting review",       "Public form → App Inbox"],
    ["📚 Submissions — Vault",      "Vault submissions awaiting review",      "Public form → App Inbox"],
    ["📅 Submissions — Events",     "Event submissions awaiting review",      "Public form → App Inbox"],
    ["", "", ""],
    ["People & Contacts",   "Approved / direct people entries",     "Direct saves + Trusted PIN"],
    ["Networking",          "Approved / direct networking entries", "Direct saves + Trusted PIN"],
    ["Shows & Events",      "Approved / direct show entries",       "Direct saves + Trusted PIN"],
    ["Expos & Markets",     "Approved / direct expo entries",       "Direct saves + Trusted PIN"],
    ["Modalities",          "Approved / direct modality entries",   "Direct saves + Trusted PIN"],
    ["Stores & Vendors",    "Approved / direct store entries",      "Direct saves + Trusted PIN"],
    ["Useful Links",        "Approved / direct link entries",       "Direct saves + Trusted PIN"],
    ["Ideas & Experiments", "Approved / direct idea entries",       "Direct saves + Trusted PIN"],
    ["Resource Vault",      "Approved / direct vault entries",      "Direct saves + Trusted PIN"],
    ["Upcoming Events",     "Approved / direct event submissions",  "Direct saves + Trusted PIN"],
  ];

  const startRow = 5;
  sections.forEach((row, i) => {
    const r = sheet.getRange(startRow + i, 1, 1, 3);
    r.setValues([row]);
    if (i === 0) r.setFontWeight("bold").setBackground("#f0eee1");
  });

  const instrRow = startRow + sections.length + 2;
  sheet.getRange(instrRow, 1).setValue("── HOW TO CONNECT ───────────────────────────────────────────────────────");
  sheet.getRange(instrRow, 1).setFontWeight("bold").setFontColor("#6b2fa0");

  [
    "1.  Fill in TRUSTED_PIN and NOTIFY_EMAIL at the top of this script",
    "",
    "2.  Already deployed! Your Web App URL:",
    "      https://script.google.com/macros/s/AKfycbzg5yEU77kQPmY3x1qbFORpwPqco9C3Oghj-vE7DN2NUbY7WTz-Vz8IPkW5Jg9Sm1o/exec",
    "      To update after edits: Deploy → New Deployment → Web App",
    "",
    "3.  In the SoulSpark app, click 'Connect Google Sheets'",
    "      Paste the URL above → Test & Connect → ✓ Connected",
    "",
    "4.  Set your Admin passcode in the SoulSpark app",
    "      Unlocks the 📥 Inbox tab for approving/rejecting submissions",
    "",
    "── REVIEW WORKFLOW ───────────────────────────────────────────────────────",
    "",
    "  • Public submits via the SoulSpark submission form → lands in Submissions tab",
    "  • Open SoulSpark app → Admin → unlock → 📥 Inbox",
    "  • Click ✓ Approve & Import → entry goes live in the guide instantly",
    "  • Or click ✕ Reject / 🔁 Needs Info → status updates in this sheet",
    "",
    "── STATUS COLORS ─────────────────────────────────────────────────────────",
    "",
    "  🟡 Amber  = ⏳ Pending Review",
    "  🟢 Green  = ✅ Approved",
    "  🔴 Red    = ❌ Rejected",
    "  🔵 Blue   = 🔁 Needs More Info",
  ].forEach((line, i) => {
    sheet.getRange(instrRow + 1 + i, 1).setValue(line).setFontSize(10);
  });

  sheet.setColumnWidth(1, 340);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 200);

  return true;
}


// ══════════════════════════════════════════════════════════════
//  UTILITIES
// ══════════════════════════════════════════════════════════════

// Convert a header string to snake_case key
//   "Submitted At"  →  "submitted_at"
//   "Cat — Books"   →  "cat_books"
function headerToSnake(h) {
  return String(h).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

// Save a base64 image to a shared Google Drive folder
function saveImageToDrive(b64, mimeType, filename, entryTitle) {
  const FOLDER_NAME = "SoulSpark Submissions — Images";
  let folder;
  const iter = DriveApp.getFoldersByName(FOLDER_NAME);
  if (iter.hasNext()) {
    folder = iter.next();
  } else {
    folder = DriveApp.createFolder(FOLDER_NAME);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  const bytes = Utilities.base64Decode(b64);
  const blob  = Utilities.newBlob(bytes, mimeType, filename);
  const file  = folder.createFile(blob);
  file.setName(entryTitle ? entryTitle + " — " + filename : filename);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// Send a notification email to the team on each new submission
function notifyTeam(sec, data, isTrusted) {
  if (!NOTIFY_EMAIL) return;
  try {
    const submitter    = data.submitter_name || "Anonymous";
    const entry        = data.name || data.org_name || data.event_name || data.expo_name ||
                         data.modality_name || data.store_name || data.link_name ||
                         data.idea_name || data.title || "—";
    const sectionLabel = isTrusted ? (MAIN_TABS[sec] || sec) : (SUBMISSION_TABS[sec] || sec);
    const trustedNote  = isTrusted ? " [TRUSTED — auto-approved]" : "";
    const action       = isTrusted
      ? "This entry was auto-approved via Trusted PIN and written directly to the main guide."
      : "Open the SoulSpark Inbox tab (Admin → 📥 Inbox) to approve or reject it.";

    MailApp.sendEmail({
      to:      NOTIFY_EMAIL,
      subject: "✦ New SoulSpark Submission" + trustedNote + " — " + sectionLabel,
      body: [
        "A new entry was submitted to your SoulSpark Resource Guide.",
        "",
        "Section      : " + sectionLabel,
        "Entry        : " + entry,
        "Submitted by : " + submitter,
        "Trusted PIN  : " + (isTrusted ? "Yes" : "No — needs review"),
        "Time         : " + new Date().toLocaleString("en-US"),
        "",
        action,
        "",
        "— SoulSpark Submission System",
      ].join("\n"),
    });
  } catch (mailErr) {
    console.log("Email notification failed: " + mailErr.message);
  }
}

// JSON response wrapper for doGet / doPost
function respond(obj, callbackParam) {
  const json = JSON.stringify(obj);
  // JSONP support — if ?callback=fnName is passed, wrap in fnName(...)
  // This lets plain HTML files fetch data without CORS errors.
  if (callbackParam) {
    return ContentService
      .createTextOutput(callbackParam + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════════════
//  OPTIONAL UTILITIES  (run manually from the editor if needed)
// ══════════════════════════════════════════════════════════════

// Wipe and rebuild a single submission tab (use with caution — deletes data).
function resetSubmissionTab(sec) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = SUBMISSION_TABS[sec];
  if (!tabName) { Logger.log("Unknown section: " + sec); return; }
  const existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);
  buildTab(ss, tabName, sec, true);
  Logger.log("Reset complete: " + tabName);
}

// Print all sheet names to the log for debugging.
function listSheets() {
  const names = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
  Logger.log(names.join("\n"));
}
