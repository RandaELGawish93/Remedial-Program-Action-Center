/**
 * EMAILFUNCTIONS.gs
 * Every function here sends through sendLoggedEmail_() (Setup_EmailLog.gs),
 * so every send is automatically recorded in Email Log — nothing here
 * calls MailApp directly.
 *
 * STYLE: emailWrapper_() below provides one shared HTML layout so every
 * email this system sends looks consistent. Edit it to match your branding.
 */

function emailWrapper_(heading, bodyHtml) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">
      <div style="background:#1c4587;color:#fff;padding:14px 20px;border-radius:6px 6px 0 0;">
        <strong style="font-size:16px;">${heading}</strong>
      </div>
      <div style="border:1px solid #e0e0e0;border-top:none;padding:20px;border-radius:0 0 6px 6px;">
        ${bodyHtml}
        <p style="margin-top:24px;font-size:12px;color:#888;">
          Office of the Academic Dean
        </p>
      </div>
    </div>`;
}

// ---------- 1. Missed Action Plan / Progress Report → Teacher ----------

// ---------- 2. Remedial Session Punctuality — Thank You / Late / Absent ----------

