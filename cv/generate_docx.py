#!/usr/bin/env python3
"""Generate an ATS-friendly single-column DOCX from the Laravel CV content."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Twips
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


OUT = Path(__file__).with_name("Muhammad_Hyder_Ali_Laravel_Developer.docx")


def set_run_font(run, size=11, bold=False):
    run.bold = bold
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), "Calibri")
    rFonts.set(qn("w:hAnsi"), "Calibri")
    rFonts.set(qn("w:eastAsia"), "Calibri")


def set_paragraph_spacing(paragraph, before=0, after=4, line=240):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line / 240


def add_horizontal_line(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "666666")
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_heading_line(doc, text):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=10, after=4)
    run = p.add_run(text.upper())
    set_run_font(run, size=12, bold=True)
    add_horizontal_line(p)
    return p


def add_body(doc, text, bold=False, size=11, before=0, after=4):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=before, after=after)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    set_paragraph_spacing(p, before=0, after=2)
    # Clear default run and add styled run
    if p.runs:
        p.runs[0].text = text
        set_run_font(p.runs[0], size=10.5)
    else:
        run = p.add_run(text)
        set_run_font(run, size=10.5)
    return p


def add_skill_line(doc, label, value):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=2)
    r1 = p.add_run(f"{label}: ")
    set_run_font(r1, size=10.5, bold=True)
    r2 = p.add_run(value)
    set_run_font(r2, size=10.5)
    return p


def main():
    doc = Document()

    # Narrow margins for 1-page ATS layout
    for section in doc.sections:
        section.top_margin = Twips(720)  # 0.5"
        section.bottom_margin = Twips(720)
        section.left_margin = Twips(720)
        section.right_margin = Twips(720)

    # Name
    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(name, before=0, after=2)
    r = name.add_run("MUHAMMAD HYDER ALI")
    set_run_font(r, size=16, bold=True)

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(title, before=0, after=2)
    r = title.add_run("Laravel Developer | PHP 8 | PostgreSQL | RESTful APIs | Laravel Cloud")
    set_run_font(r, size=11, bold=False)

    # Contact
    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(contact, before=0, after=2)
    r = contact.add_run(
        "Multan, Pakistan (Remote · UK timezone overlap) | +92 313 6590599 | hyderali11112@gmail.com"
    )
    set_run_font(r, size=10)

    links = doc.add_paragraph()
    links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(links, before=0, after=6)
    r = links.add_run(
        "LinkedIn: linkedin.com/in/muhammadhyderali | GitHub: github.com/MuhammadHyderAli"
    )
    set_run_font(r, size=10)

    # Summary
    add_heading_line(doc, "Professional Summary")
    add_body(
        doc,
        "Laravel Developer with 2+ years building production web applications and APIs in PHP 8 and Laravel. "
        "Owns features end-to-end—from Filament/Livewire admin UIs through RESTful backends, PostgreSQL/MySQL data models, "
        "and multi-system integrations. Deployed and maintained commercial SaaS on Laravel Cloud. Uses Claude and Cursor "
        "as disciplined engineering tools to ship faster without sacrificing code quality, reviews, or ownership.",
        size=10.5,
        after=4,
    )

    # Skills
    add_heading_line(doc, "Technical Skills")
    add_skill_line(
        doc,
        "Backend",
        "PHP 8+, Laravel (10, 11), Eloquent ORM, Queues, Jobs, Events & Listeners, Middleware, Form Requests, Sanctum, Policies & Gates, Multi-Tenancy, RBAC",
    )
    add_skill_line(
        doc,
        "Databases",
        "PostgreSQL, MySQL, Relational Design, Indexing, Transactions, Complex Joins",
    )
    add_skill_line(
        doc,
        "APIs & Integrations",
        "RESTful API Design, JSON, CORS, API Rate Limiting, Xero, DocuSign, Twilio, OpenAI, Google Gemini",
    )
    add_skill_line(
        doc,
        "User-Facing Web",
        "Laravel Filament, Livewire, Vue.js, React.js, Next.js, Tailwind CSS",
    )
    add_skill_line(
        doc,
        "Cloud & Quality",
        "Laravel Cloud, AWS S3, Laravel Cache, N+1 Resolution, PHPUnit, Pest, Git, Agile",
    )
    add_skill_line(
        doc,
        "AI Tooling",
        "Claude, Cursor — AI-assisted development for planning, implementation, refactoring and debugging",
    )

    # Experience
    add_heading_line(doc, "Professional Experience")

    role = doc.add_paragraph()
    set_paragraph_spacing(role, before=2, after=0)
    r = role.add_run("Associate Software Engineer (Backend — PHP / Laravel) | Propday (HighGate / Remax)")
    set_run_font(r, size=10.5, bold=True)

    meta = doc.add_paragraph()
    set_paragraph_spacing(meta, before=0, after=2)
    r = meta.add_run("Multan, Pakistan | Jun 2025 – Present")
    set_run_font(r, size=10)

    propday_bullets = [
        "Own backend delivery for a production multi-tenant property management SaaS in PHP 8 and Laravel, consolidating external platforms into one central system.",
        "Led migration from a Filament monolith to an API-first Laravel backend with a decoupled Next.js frontend, owning features from UI admin workflows through to API and database design.",
        "Designed RBAC and multi-tenancy across Super Admin, Company and User hierarchies with Laravel policies, gates and tenant data isolation.",
        "Built core modules with Eloquent, migrations and database transactions: financial ledgers, bank statements, work orders, applicants and property management on relational SQL (PostgreSQL / MySQL).",
        "Integrated third-party APIs (including Xero) and implemented queues, jobs, events, notifications and scheduled tasks for reliable syncs and automated workflows.",
        "Deploy and maintain the application on Laravel Cloud; improved performance with eager loading and Laravel Cache; store files on AWS S3; write PHPUnit tests and conduct code reviews.",
    ]
    for b in propday_bullets:
        add_bullet(doc, b)

    role2 = doc.add_paragraph()
    set_paragraph_spacing(role2, before=6, after=0)
    r = role2.add_run("Full Stack Developer (Laravel) | Cyberify")
    set_run_font(r, size=10.5, bold=True)

    meta2 = doc.add_paragraph()
    set_paragraph_spacing(meta2, before=0, after=2)
    r = meta2.add_run("Multan, Pakistan | Sep 2024 – Apr 2025")
    set_run_font(r, size=10)

    cyberify_bullets = [
        "Built user-facing web and desktop clients with Laravel, Vue.js and Electron, owning features from planning through deployment.",
        "Implemented secure RESTful APIs with Laravel Sanctum, CORS and validated CRUD endpoints consumed by web and desktop clients.",
        "Delivered features in agile cycles, collaborating with frontend developers and product stakeholders.",
    ]
    for b in cyberify_bullets:
        add_bullet(doc, b)

    # Projects
    add_heading_line(doc, "Selected Projects")

    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=2, after=0)
    r = p.add_run("Mobile Application — Expo / React Native with Laravel API")
    set_run_font(r, size=10.5, bold=True)
    add_body(
        doc,
        "Cross-platform Android/iOS app backed by a Laravel REST API with Sanctum token authentication and validated JSON endpoints. Built both the mobile client and the Laravel API it consumes.",
        size=10.5,
        after=4,
    )

    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=2, after=0)
    r = p.add_run("AI Voice Agent — Real-Time Telephony & LLM Backend")
    set_run_font(r, size=10.5, bold=True)
    add_body(
        doc,
        "Laravel backend for a real-time AI voice agent: Twilio telephony webhooks feed a call-context pipeline that routes prompts to OpenAI / Gemini and returns responses to the caller. Demonstrates disciplined use of AI APIs inside production Laravel infrastructure.",
        size=10.5,
        after=4,
    )

    # Education
    add_heading_line(doc, "Education")
    edu = doc.add_paragraph()
    set_paragraph_spacing(edu, before=2, after=0)
    r = edu.add_run("BS Software Engineering")
    set_run_font(r, size=10.5, bold=True)
    r = edu.add_run(" | FAST University (CFD Campus) | 2019 – 2024")
    set_run_font(r, size=10.5)

    doc.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
