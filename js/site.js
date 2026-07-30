/* ==========================================================================
   EAT COE — application logic
   No signup, no backend: everything below runs entirely client-side.
   Real document FILES always live in SharePoint/Teams — this site only
   stores a catalog entry (name, tags, owner, link). "Register a document"
   never uploads a file anywhere; it just links out to SharePoint.
   Anything a visitor adds/edits (registrations, tags, status, role) is
   saved in that visitor's own browser only (localStorage) — it is not
   shared with other people opening the same link. See README.md for how
   to make the catalog shared across everyone using a free backend.
   ========================================================================== */

const SHAREPOINT_FOLDER_URL = "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP";

const STORIES = [
  {
    "code": "1.1",
    "title": "Define Test Strategy, Compliance & SOPs",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "status": "Done",
    "owner": "QA Lead / Manager",
    "description": "Unified test strategy & SOPs, plus the RAT/RTM template set.",
    "tags": [
      "strategy",
      "sop",
      "compliance",
      "rtm",
      "test plan"
    ],
    "url": "01-standards.html#1.1",
    "doc": "docs/01-standards/1.1 Define Test strategy, compliance, and SOPs/Unified Test Strategy & SOPs for ERP.docx",
    "type": "story"
  },
  {
    "code": "1.2",
    "title": "Reusable Automation Framework with CI/CD",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "status": "Done",
    "owner": "Automation Architect",
    "description": "Automation framework architecture, reusable function library, data-driven test design guide.",
    "tags": [
      "automation",
      "framework",
      "ci/cd",
      "architecture"
    ],
    "url": "01-standards.html#1.2",
    "type": "story"
  },
  {
    "code": "1.3",
    "title": "Quality Metrics & KPIs",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "status": "Done",
    "owner": "QA Manager",
    "description": "KPI catalog (e.g. DRE, cycle time) and dashboard specification.",
    "tags": [
      "kpi",
      "metrics",
      "dashboard",
      "dre"
    ],
    "url": "01-standards.html#1.3",
    "doc": "docs/01-standards/1.3 Quality metrics & KPIs/ERP Quality Metrics & KPIs.docx",
    "type": "story"
  },
  {
    "code": "1.4",
    "title": "Monitoring & Governance Framework",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "status": "Done",
    "owner": "COE Lead",
    "description": "Governance playbook, review calendar, escalation matrix, governance audit checklist.",
    "tags": [
      "governance",
      "monitoring",
      "escalation",
      "playbook"
    ],
    "url": "01-standards.html#1.4",
    "doc": "docs/01-standards/1.4 Monitoring & Governance framework/Monitoring & Governance Framework for ERP.docx",
    "type": "story"
  },
  {
    "code": "1.5",
    "title": "Versioning & Change Control Process",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "status": "Done",
    "owner": "COE Lead",
    "description": "Versioning approach, change log template, adoption training.",
    "tags": [
      "versioning",
      "change control",
      "change log"
    ],
    "url": "01-standards.html#1.5",
    "doc": "docs/01-standards/1.5 Versioning and Change Control Process/Versioning and Change Control Process for ERP.docx",
    "type": "story"
  },
  {
    "code": "1.6",
    "title": "Test Data Management Standards",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "status": "In Progress",
    "owner": "ETL Lead",
    "description": "TDM strategy per ERP module, TDM SOP, anonymization/masking guidelines.",
    "tags": [
      "test data",
      "tdm",
      "masking",
      "anonymization"
    ],
    "url": "01-standards.html#1.6",
    "type": "story"
  },
  {
    "code": "2.1",
    "title": "Tool Evaluation Framework",
    "pillarCode": "02",
    "pillar": "Tools & Technology",
    "status": "Done",
    "owner": "Automation Architect",
    "description": "Scoring criteria (cost, support, flexibility, ROI) and evaluation template.",
    "tags": [
      "tool evaluation",
      "scoring",
      "roi"
    ],
    "url": "02-tools.html#2.1",
    "type": "story"
  },
  {
    "code": "2.2",
    "title": "Tool Capability & Best Practices",
    "pillarCode": "02",
    "pillar": "Tools & Technology",
    "status": "In Progress",
    "owner": "QA Lead",
    "description": "Tool documentation library with worked use-case examples.",
    "tags": [
      "tools",
      "documentation",
      "best practices"
    ],
    "url": "02-tools.html#2.2",
    "type": "story"
  },
  {
    "code": "2.3",
    "title": "Checklists & User Guides",
    "pillarCode": "02",
    "pillar": "Tools & Technology",
    "status": "Done",
    "owner": "QA Lead",
    "description": "Role-based guides with screenshots and tips.",
    "tags": [
      "checklist",
      "user guide",
      "how to"
    ],
    "url": "02-tools.html#2.3",
    "type": "story"
  },
  {
    "code": "2.4",
    "title": "Integration with CI/CD Tools",
    "pillarCode": "02",
    "pillar": "Tools & Technology",
    "status": "Backlog",
    "owner": "Automation Architect",
    "description": "CI/CD integration standards, reusable pipeline scripts, use-case library (Jenkins, Azure DevOps, GitLab).",
    "tags": [
      "ci/cd",
      "jenkins",
      "azure devops",
      "gitlab",
      "pipeline"
    ],
    "url": "02-tools.html#2.4",
    "type": "story"
  },
  {
    "code": "3.1",
    "title": "Market Study & Trend Analysis",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "status": "In Progress",
    "owner": "COE Member",
    "description": "Quarterly trend reports with recommendations to the COE.",
    "tags": [
      "trends",
      "market study",
      "ai/ml"
    ],
    "url": "03-monitoring.html#3.1",
    "type": "story"
  },
  {
    "code": "3.2",
    "title": "Lessons Learnt",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "status": "Backlog",
    "owner": "QA Lead",
    "description": "Lessons-learned template and post-release retrospective log.",
    "tags": [
      "lessons learned",
      "retrospective"
    ],
    "url": "03-monitoring.html#3.2",
    "type": "story"
  },
  {
    "code": "3.3",
    "title": "Review Process",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "status": "Backlog",
    "owner": "QA Lead",
    "description": "Review checklist and mandatory pre-UAT review gate.",
    "tags": [
      "review",
      "checklist",
      "uat"
    ],
    "url": "03-monitoring.html#3.3",
    "type": "story"
  },
  {
    "code": "3.4",
    "title": "Monitoring & Governance Practices",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "status": "To Do",
    "owner": "QA Manager",
    "description": "Monitoring SOPs, DSR/WSR/dashboard templates, alerting configuration.",
    "tags": [
      "monitoring",
      "dsr",
      "wsr",
      "alerting"
    ],
    "url": "03-monitoring.html#3.4",
    "type": "story"
  },
  {
    "code": "3.5",
    "title": "Continuous Improvement",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "status": "Done",
    "owner": "COE Lead",
    "description": "Continuous improvement framework and tracked action log.",
    "tags": [
      "continuous improvement",
      "action log"
    ],
    "url": "03-monitoring.html#3.5",
    "doc": "docs/03-monitoring/3.5 Continuous Improvement/Continuous Improvement Framework for ERP.docx",
    "type": "story"
  },
  {
    "code": "3.6",
    "title": "Defect Trend Analysis",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "status": "Done",
    "owner": "QA Lead",
    "description": "Defect trend methodology, dashboard templates, RCA format.",
    "tags": [
      "defects",
      "trend analysis",
      "rca",
      "dashboard"
    ],
    "url": "03-monitoring.html#3.6",
    "doc": "docs/03-monitoring/3.6 Defect Trend Analysis/Defect Trend Analysis and Reporting Framework for ERP Testing.docx",
    "type": "story"
  },
  {
    "code": "4.1",
    "title": "Success & Failure Reports (Case Studies)",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "Backlog",
    "owner": "COE Member",
    "description": "Quarterly case study, reviewed and published internally.",
    "tags": [
      "case study",
      "success",
      "failure"
    ],
    "url": "04-capability.html#4.1",
    "type": "story"
  },
  {
    "code": "4.2",
    "title": "Capability Deck",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "To Do",
    "owner": "QA Lead",
    "description": "Capability deck, refreshed quarterly.",
    "tags": [
      "capability deck",
      "onboarding"
    ],
    "url": "04-capability.html#4.2",
    "type": "story"
  },
  {
    "code": "4.3",
    "title": "Skill Development",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "To Do",
    "owner": "COE Member",
    "description": "Skill matrix, learning paths, trainer roster, training calendar.",
    "tags": [
      "skills",
      "training",
      "learning path"
    ],
    "url": "04-capability.html#4.3",
    "type": "story"
  },
  {
    "code": "4.4",
    "title": "Knowledge Management Framework",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "Done",
    "owner": "COE Lead",
    "description": "Central repository with indexing and folder structure — the foundation this website implements.",
    "tags": [
      "knowledge management",
      "repository",
      "indexing"
    ],
    "url": "04-capability.html#4.4",
    "doc": "docs/04-capability/4.4 Knowledge Management Framework/Knowledge Management Framework for ERP Testing.docx",
    "type": "story"
  },
  {
    "code": "4.5",
    "title": "Business & Domain Knowledge",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "Done",
    "owner": "COE Member",
    "description": "Module synopsis pages indexed with links to related documents, SME walkthroughs.",
    "tags": [
      "business knowledge",
      "domain knowledge",
      "sme"
    ],
    "url": "04-capability.html#4.5",
    "doc": "docs/04-capability/4.5 Business & Domain knowledge/Business & Domain Knowledge Framework for ERP.docx",
    "type": "story"
  },
  {
    "code": "4.6",
    "title": "Business / Domain SMEs",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "In Progress",
    "owner": "QA Director",
    "description": "SME directory by module with a monthly sync cadence.",
    "tags": [
      "sme",
      "directory",
      "domain experts"
    ],
    "url": "04-capability.html#4.6",
    "type": "story"
  },
  {
    "code": "4.7",
    "title": "Career Growth & Mentoring",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "To Do",
    "owner": "COE Lead",
    "description": "Training curriculum with notes/presentations, attendance and feedback tracking.",
    "tags": [
      "mentoring",
      "career growth",
      "curriculum"
    ],
    "url": "04-capability.html#4.7",
    "type": "story"
  },
  {
    "code": "4.8",
    "title": "ERP-Specific Scenarios & Use Cases",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "status": "Done",
    "owner": "COE Member",
    "description": "Business/functionality-specific standard scenario library, reviewed quarterly.",
    "tags": [
      "scenarios",
      "use cases",
      "erp"
    ],
    "url": "04-capability.html#4.8",
    "type": "story"
  },
  {
    "code": "5.1",
    "title": "Checklists & Templates",
    "pillarCode": "05",
    "pillar": "Innovation",
    "status": "In Progress",
    "owner": "QA Lead",
    "description": "Published starter kits with documented usage.",
    "tags": [
      "ai",
      "checklist",
      "template",
      "starter kit"
    ],
    "url": "05-innovation.html#5.1",
    "type": "story"
  },
  {
    "code": "5.2",
    "title": "AI-Assisted Testing Workflows",
    "pillarCode": "05",
    "pillar": "Innovation",
    "status": "Backlog",
    "owner": "QA Lead",
    "description": "Workflow guidelines with a risk log, rollback plan and review checkpoints.",
    "tags": [
      "ai",
      "workflow",
      "risk log",
      "rollback"
    ],
    "url": "05-innovation.html#5.2",
    "type": "story"
  },
  {
    "code": "5.3",
    "title": "AI Test Optimization",
    "pillarCode": "05",
    "pillar": "Innovation",
    "status": "Backlog",
    "owner": "Automation Engineer",
    "description": "Proof of concept with pilot results shared.",
    "tags": [
      "ai",
      "optimization",
      "poc",
      "pilot"
    ],
    "url": "05-innovation.html#5.3",
    "type": "story"
  },
  {
    "code": "5.4",
    "title": "Develop Accelerators",
    "pillarCode": "05",
    "pillar": "Innovation",
    "status": "Backlog",
    "owner": "QA Lead",
    "description": "Accelerator library with documented usage.",
    "tags": [
      "accelerator",
      "data validator",
      "script generator"
    ],
    "url": "05-innovation.html#5.4",
    "type": "story"
  },
  {
    "code": "5.5",
    "title": "Implement Relevant AI Use Cases",
    "pillarCode": "05",
    "pillar": "Innovation",
    "status": "Backlog",
    "owner": "COE Lead",
    "description": "At least one production AI use case with ROI tracked post-implementation.",
    "tags": [
      "ai",
      "use case",
      "roi",
      "production"
    ],
    "url": "05-innovation.html#5.5",
    "type": "story"
  }
];

const SEED_DOCS = [
  {
    "id": "seed-1.1",
    "name": "Unified Test Strategy & SOPs for ERP.docx",
    "storyCode": "1.1",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "uploadedBy": "QA Manager",
    "uploadDate": "2026-02-10",
    "lastModifiedBy": "QA Manager",
    "lastModifiedDate": "2026-03-18",
    "tags": [
      "strategy",
      "sop",
      "compliance",
      "rtm"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 58,
    "location": "SharePoint · DigitalBanking",
    "featured": true,
    "type": "document",
    "sourceType": "seed"
  },
  {
    "id": "seed-1.3",
    "name": "ERP Quality Metrics & KPIs.docx",
    "storyCode": "1.3",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "uploadedBy": "QA Manager",
    "uploadDate": "2026-01-22",
    "lastModifiedBy": "QA Manager",
    "lastModifiedDate": "2026-01-22",
    "tags": [
      "kpi",
      "metrics",
      "dashboard"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 41,
    "location": "SharePoint · DigitalBanking",
    "featured": false,
    "type": "document",
    "sourceType": "seed"
  },
  {
    "id": "seed-1.4",
    "name": "Monitoring & Governance Framework for ERP.docx",
    "storyCode": "1.4",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "uploadedBy": "COE Lead",
    "uploadDate": "2025-12-05",
    "lastModifiedBy": "COE Lead",
    "lastModifiedDate": "2026-04-02",
    "tags": [
      "governance",
      "monitoring",
      "escalation"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 33,
    "location": "SharePoint · DigitalBanking",
    "featured": false,
    "type": "document",
    "sourceType": "seed"
  },
  {
    "id": "seed-1.5",
    "name": "Versioning and Change Control Process for ERP.docx",
    "storyCode": "1.5",
    "pillarCode": "01",
    "pillar": "Standards & Best Practices",
    "uploadedBy": "COE Lead",
    "uploadDate": "2025-11-14",
    "lastModifiedBy": "COE Lead",
    "lastModifiedDate": "2025-11-14",
    "tags": [
      "versioning",
      "change control"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 19,
    "location": "SharePoint · DigitalBanking",
    "featured": false,
    "type": "document",
    "sourceType": "seed"
  },
  {
    "id": "seed-3.5",
    "name": "Continuous Improvement Framework for ERP.docx",
    "storyCode": "3.5",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "uploadedBy": "COE Lead",
    "uploadDate": "2026-03-01",
    "lastModifiedBy": "COE Lead",
    "lastModifiedDate": "2026-05-11",
    "tags": [
      "continuous improvement",
      "action log"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 47,
    "location": "SharePoint · DigitalBanking",
    "featured": false,
    "type": "document",
    "sourceType": "seed"
  },
  {
    "id": "seed-3.6",
    "name": "Defect Trend Analysis and Reporting Framework for ERP Testing.docx",
    "storyCode": "3.6",
    "pillarCode": "03",
    "pillar": "Monitoring & Continuous Improvement",
    "uploadedBy": "QA Lead",
    "uploadDate": "2026-04-19",
    "lastModifiedBy": "QA Lead",
    "lastModifiedDate": "2026-06-30",
    "tags": [
      "defects",
      "trend analysis",
      "rca"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 62,
    "location": "SharePoint · DigitalBanking",
    "featured": true,
    "type": "document",
    "sourceType": "seed"
  },
  {
    "id": "seed-4.4",
    "name": "Knowledge Management Framework for ERP Testing.docx",
    "storyCode": "4.4",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "uploadedBy": "COE Lead",
    "uploadDate": "2025-10-08",
    "lastModifiedBy": "COE Lead",
    "lastModifiedDate": "2026-02-14",
    "tags": [
      "knowledge management",
      "repository"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 29,
    "location": "SharePoint · DigitalBanking",
    "featured": true,
    "type": "document",
    "sourceType": "seed"
  },
  {
    "id": "seed-4.5",
    "name": "Business & Domain Knowledge Framework for ERP.docx",
    "storyCode": "4.5",
    "pillarCode": "04",
    "pillar": "Capability Building & Enablement",
    "uploadedBy": "COE Member",
    "uploadDate": "2026-05-27",
    "lastModifiedBy": "COE Member",
    "lastModifiedDate": "2026-07-05",
    "tags": [
      "business knowledge",
      "domain knowledge",
      "sme"
    ],
    "url": "https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84?e=3udilP",
    "downloads": 22,
    "location": "SharePoint · DigitalBanking",
    "featured": false,
    "type": "document",
    "sourceType": "seed"
  }
];

const SEED_ACTIVITY = [
  {
    "type": "download",
    "actor": "COE Lead",
    "target": "Business & Domain Knowledge Framework for ERP.docx",
    "timestamp": "2026-07-16"
  },
  {
    "type": "status",
    "actor": "QA Manager",
    "target": "a story status",
    "timestamp": "2026-07-17"
  },
  {
    "type": "upload",
    "actor": "Automation Architect",
    "target": "Versioning and Change Control Process for ERP.docx",
    "timestamp": "2026-07-18"
  },
  {
    "type": "download",
    "actor": "Automation Architect",
    "target": "Monitoring & Governance Framework for ERP.docx",
    "timestamp": "2026-07-18"
  },
  {
    "type": "upload",
    "actor": "Automation Architect",
    "target": "Versioning and Change Control Process for ERP.docx",
    "timestamp": "2026-07-18"
  },
  {
    "type": "tag",
    "actor": "QA Manager",
    "target": "ERP Quality Metrics & KPIs.docx",
    "timestamp": "2026-07-18"
  },
  {
    "type": "status",
    "actor": "COE Member",
    "target": "a story status",
    "timestamp": "2026-07-19"
  },
  {
    "type": "status",
    "actor": "COE Member",
    "target": "a story status",
    "timestamp": "2026-07-19"
  },
  {
    "type": "upload",
    "actor": "QA Manager",
    "target": "Defect Trend Analysis and Reporting Framework for ERP Testing.docx",
    "timestamp": "2026-07-19"
  },
  {
    "type": "download",
    "actor": "COE Member",
    "target": "Unified Test Strategy & SOPs for ERP.docx",
    "timestamp": "2026-07-20"
  },
  {
    "type": "tag",
    "actor": "COE Lead",
    "target": "ERP Quality Metrics & KPIs.docx",
    "timestamp": "2026-07-20"
  },
  {
    "type": "upload",
    "actor": "Automation Architect",
    "target": "ERP Quality Metrics & KPIs.docx",
    "timestamp": "2026-07-20"
  },
  {
    "type": "download",
    "actor": "COE Lead",
    "target": "Versioning and Change Control Process for ERP.docx",
    "timestamp": "2026-07-20"
  },
  {
    "type": "tag",
    "actor": "QA Lead",
    "target": "ERP Quality Metrics & KPIs.docx",
    "timestamp": "2026-07-20"
  },
  {
    "type": "download",
    "actor": "Automation Architect",
    "target": "Business & Domain Knowledge Framework for ERP.docx",
    "timestamp": "2026-07-20"
  },
  {
    "type": "tag",
    "actor": "QA Manager",
    "target": "Unified Test Strategy & SOPs for ERP.docx",
    "timestamp": "2026-07-20"
  },
  {
    "type": "status",
    "actor": "QA Manager",
    "target": "a story status",
    "timestamp": "2026-07-22"
  },
  {
    "type": "upload",
    "actor": "COE Member",
    "target": "Knowledge Management Framework for ERP Testing.docx",
    "timestamp": "2026-07-22"
  },
  {
    "type": "download",
    "actor": "QA Manager",
    "target": "Knowledge Management Framework for ERP Testing.docx",
    "timestamp": "2026-07-22"
  },
  {
    "type": "status",
    "actor": "QA Manager",
    "target": "a story status",
    "timestamp": "2026-07-22"
  },
  {
    "type": "download",
    "actor": "QA Manager",
    "target": "Monitoring & Governance Framework for ERP.docx",
    "timestamp": "2026-07-22"
  },
  {
    "type": "status",
    "actor": "Automation Architect",
    "target": "a story status",
    "timestamp": "2026-07-22"
  },
  {
    "type": "status",
    "actor": "QA Lead",
    "target": "a story status",
    "timestamp": "2026-07-22"
  },
  {
    "type": "status",
    "actor": "COE Member",
    "target": "a story status",
    "timestamp": "2026-07-22"
  }
];

/* ---------------- localStorage helpers ---------------- */
const LS = {
  role: "eatcoe_role",
  user: "eatcoe_username",
  userDocs: "eatcoe_user_docs",
  seedOverrides: "eatcoe_seed_overrides",
  statusOverrides: "eatcoe_status_overrides",
  activity: "eatcoe_activity",
  downloads: "eatcoe_downloads"
};

function lsGet(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw === null ? fallback : JSON.parse(raw);
  } catch (e) { return fallback; }
}
function lsSet(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { console.error("storage failed", e); }
}

function getRole() { return lsGet(LS.role, "viewer"); }
function setRole(r) { lsSet(LS.role, r); applyRoleToBody(); if (typeof syncStatusEditors === "function") syncStatusEditors(); }
function getUserName() { return lsGet(LS.user, "Guest"); }
function setUserName(n) { lsSet(LS.user, n || "Guest"); }

function applyRoleToBody() {
  document.body.classList.toggle("role-viewer", getRole() !== "contributor");
  document.body.classList.toggle("role-contributor", getRole() === "contributor");
}

function initRoleSwitcher() {
  applyRoleToBody();
}

/* ---------------- activity log ---------------- */
function ensureSeedActivity() {
  const existing = localStorage.getItem(LS.activity);
  if (existing === null) lsSet(LS.activity, SEED_ACTIVITY);
}
function getActivity() { return lsGet(LS.activity, SEED_ACTIVITY); }
function logActivity(type, target) {
  const log = getActivity();
  log.push({ type, actor: getUserName(), target, timestamp: new Date().toISOString().slice(0, 10) });
  lsSet(LS.activity, log.slice(-500));
}

/* ---------------- documents ---------------- */
function getUserDocs() { return lsGet(LS.userDocs, []); }
function saveUserDocs(arr) { lsSet(LS.userDocs, arr); }
function getSeedOverrides() { return lsGet(LS.seedOverrides, {}); }
function saveSeedOverrides(obj) { lsSet(LS.seedOverrides, obj); }
function getDownloadCounts() { return lsGet(LS.downloads, {}); }
function bumpDownload(docId) {
  const counts = getDownloadCounts();
  counts[docId] = (counts[docId] || 0) + 1;
  lsSet(LS.downloads, counts);
}

// Populated asynchronously by js/graph.js once the real SharePoint folder
// listing has been fetched via Microsoft Graph. Until then (not signed in,
// Graph not configured, or still loading), getAllDocuments() falls back to
// the static seed catalog below, so the site always shows something.
let GRAPH_DOCS = null;
function setGraphDocuments(list) {
  GRAPH_DOCS = list;
  refreshDocViews();
  if (typeof renderMetrics === "function") renderMetrics();
}

function getAllDocuments() {
  const overrides = getSeedOverrides();
  const counts = getDownloadCounts();

  let base;
  if (GRAPH_DOCS && GRAPH_DOCS.length) {
    // Real documents fetched live from SharePoint via Microsoft Graph.
    base = GRAPH_DOCS.map(d => ({ ...d, downloads: d.downloads + (counts[d.id] || 0) }));
  } else {
    base = SEED_DOCS.map(d => {
      const o = overrides[d.id] || {};
      return {
        ...d,
        tags: o.tags || d.tags,
        lastModifiedBy: o.lastModifiedBy || d.lastModifiedBy,
        lastModifiedDate: o.lastModifiedDate || d.lastModifiedDate,
        downloads: d.downloads + (counts[d.id] || 0)
      };
    });
  }

  const user = getUserDocs().map(d => ({ ...d, downloads: d.downloads + (counts[d.id] || 0) }));
  const pending = getPendingAttachment();
  return [...base, ...user, ...(pending ? [pending] : [])];
}

function getDocsForStory(code) {
  return getAllDocuments().filter(d => d.storyCode === code);
}

/* ---------------- Register a Document (email-based contribution) ----------------
 * There is no in-browser file upload anymore: real documents live in
 * SharePoint/Teams, and files are routed there by a human who reads the
 * submitted email and files the attachment into the folder matching the
 * chosen pillar/story labels. This section only composes that email —
 * nothing here writes to the document catalog or to any storage.
 * TODO: replace with a direct Microsoft Graph upload once SSO has
 * Sites.Selected write access — at that point this form could upload for
 * real, but the mailto fallback is still worth keeping for anyone without
 * SharePoint access.
 */
const COE_INTAKE_EMAIL = "malkiat.singh@testingxperts.com"; // TODO: replace with your real intake mailbox

const PILLAR_FOLDERS = {
  "01": "docs/01-standards",
  "02": "docs/02-tools",
  "03": "docs/03-monitoring",
  "04": "docs/04-capability",
  "05": "docs/05-innovation"
};
const PILLAR_NAMES = {
  "01": "Standards & Best Practices",
  "02": "Tools & Technology",
  "03": "Monitoring & Continuous Improvement",
  "04": "Capability Building & Enablement",
  "05": "Innovation"
};

function suggestedFolder(pillarCode, storyCode) {
  if (!pillarCode || !PILLAR_FOLDERS[pillarCode]) return null;
  let path = PILLAR_FOLDERS[pillarCode];
  if (storyCode) {
    const story = STORIES.find(s => s.code === storyCode);
    if (story) path += "/" + story.code + " " + story.title;
  }
  return path + "/";
}
/* ---------------- full-text extraction (PDF.js / Mammoth.js) ----------------
 * Used in two places: (1) the file attached on Register a Document is
 * indexed for THIS browsing session only, so it's searchable even before
 * it's emailed/filed into SharePoint; (2) once Microsoft Graph is connected
 * (see js/graph.js), the same functions extract text from real SharePoint
 * files so the main search bar and chatbot can match text inside them too.
 */
const MAX_FULLTEXT_CHARS = 200000;

if (typeof pdfjsLib !== "undefined") {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3/build/pdf.worker.min.js";
}

async function extractPdfText(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(it => it.str).join(" ") + "\n";
    if (text.length > MAX_FULLTEXT_CHARS) break;
  }
  return text.slice(0, MAX_FULLTEXT_CHARS);
}

async function extractDocxText(arrayBuffer) {
  const result = await mammoth.extractRawText({ arrayBuffer });
  return (result.value || "").slice(0, MAX_FULLTEXT_CHARS);
}

/**
 * Extracts searchable full text from a File/Blob (with .name/.type/.arrayBuffer()).
 * Returns { text, status } where status is one of:
 * "ok" | "unsupported" (file type not handled) |
 * "unavailable" (PDF.js/Mammoth.js failed to load) | "error" (parse failure).
 */
async function extractFullText(file) {
  const nameLower = (file.name || "").toLowerCase();
  try {
    if (file.type === "application/pdf" || nameLower.endsWith(".pdf")) {
      if (typeof pdfjsLib === "undefined") return { text: null, status: "unavailable" };
      const buf = await file.arrayBuffer();
      return { text: await extractPdfText(buf), status: "ok" };
    }
    if (file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || nameLower.endsWith(".docx")) {
      if (typeof mammoth === "undefined") return { text: null, status: "unavailable" };
      const buf = await file.arrayBuffer();
      return { text: await extractDocxText(buf), status: "ok" };
    }
    if ((file.type || "").startsWith("text/") || nameLower.endsWith(".txt") || nameLower.endsWith(".md")) {
      const text = await file.text();
      return { text: text.slice(0, MAX_FULLTEXT_CHARS), status: "ok" };
    }
    return { text: null, status: "unsupported" };
  } catch (e) {
    console.error("Full-text extraction failed for", file.name, e);
    return { text: null, status: "error" };
  }
}

/* ---------------- pending attachment (session-only search preview) ---------------- */
const PENDING_KEY = "eatcoe_pending_attachment";
function getPendingAttachment() {
  try {
    const raw = sessionStorage.getItem(PENDING_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}
function savePendingAttachment(doc) {
  try { sessionStorage.setItem(PENDING_KEY, JSON.stringify(doc)); } catch (e) { console.error(e); }
}
function clearPendingAttachment() {
  try { sessionStorage.removeItem(PENDING_KEY); } catch (e) { /* ignore */ }
}

function initRegisterForm() {
  const form = document.getElementById("docRegisterForm");
  if (!form) return;

  const pillarSelect = document.getElementById("rfPillar");
  const storySelect = document.getElementById("rfStory");
  const fileInput = document.getElementById("rfFile");
  const fileNameLabel = document.getElementById("rfFileName");
  const nameInput = document.getElementById("rfName");
  const emailInput = document.getElementById("rfEmail");
  const docNameInput = document.getElementById("rfDocName");
  const sendBtn = document.getElementById("rfSendBtn");

  // Auto-fetch name/email from the signed-in Microsoft account (falls back
  // to the locally-stored name if SSO isn't configured yet).
  const account = (typeof getActiveAccount === "function") ? getActiveAccount() : null;
  if (account) {
    if (nameInput && !nameInput.value) nameInput.value = account.name || "";
    if (emailInput && !emailInput.value) emailInput.value = account.username || "";
  } else if (nameInput && getUserName() !== "Guest") {
    nameInput.value = getUserName();
  }

  function populateStoryOptions() {
    const pillar = pillarSelect.value;
    const current = storySelect.value;
    storySelect.innerHTML = '<option value="">No specific story</option>' +
      STORIES.filter(s => !pillar || s.pillarCode === pillar)
        .map(s => `<option value="${s.code}">${escapeHtml(s.code)} · ${escapeHtml(s.title)}</option>`).join("");
    if ([...storySelect.options].some(o => o.value === current)) storySelect.value = current;
  }
  populateStoryOptions();

  const params = new URLSearchParams(window.location.search);
  if (params.get("pillar")) pillarSelect.value = params.get("pillar");
  populateStoryOptions();
  if (params.get("story")) storySelect.value = params.get("story");

  pillarSelect.addEventListener("change", populateStoryOptions);

  fileInput.addEventListener("change", async () => {
    const file = fileInput.files[0];
    if (!file) { clearPendingAttachment(); fileNameLabel.textContent = ""; return; }

    fileNameLabel.textContent = "📎 " + file.name + " (" + Math.round(file.size / 1024) + " KB) — indexing for search…";
    if (!docNameInput.value) docNameInput.value = file.name;

    const { text, status } = await extractFullText(file);
    const today = new Date().toISOString().slice(0, 10);
    savePendingAttachment({
      id: "pending-attachment", type: "document", sourceType: "pending",
      name: docNameInput.value || file.name,
      pillarCode: pillarSelect.value || null, pillar: PILLAR_NAMES[pillarSelect.value] || null,
      storyCode: storySelect.value || null,
      tags: [], downloads: 0, featured: false,
      uploadedBy: nameInput.value || getUserName(), uploadDate: today,
      lastModifiedBy: nameInput.value || getUserName(), lastModifiedDate: today,
      url: "register-document.html",
      location: "📋 Pending submission — attached here, not yet emailed",
      fullText: text, fullTextStatus: status
    });

    fileNameLabel.textContent = "📎 " + file.name + " (" + Math.round(file.size / 1024) + " KB)" +
      (status === "ok" ? " · 🔍 indexed for search" : "");
  });

  function buildEmail() {
    const name = docNameInput.value.trim() || "(untitled document)";
    const contributor = nameInput.value.trim() || "(not provided)";
    const contributorEmail = emailInput.value.trim();
    const desc = document.getElementById("rfDescription").value.trim();
    const tags = document.getElementById("rfTags").value.trim();
    const pillar = pillarSelect.value;
    const story = storySelect.value;
    const file = fileInput.files[0];
    const folder = suggestedFolder(pillar, story);

    const subject = "TREAT COE Document Submission: " + name;
    let body = "New document submission for the TREAT COE repository.\n\n";
    body += "Document name: " + name + "\n";
    body += "Submitted by: " + contributor + (contributorEmail ? " (" + contributorEmail + ")" : "") + "\n";
    body += "Pillar: " + (PILLAR_NAMES[pillar] || "Not specified") + "\n";
    if (story) {
      const s = STORIES.find(x => x.code === story);
      body += "Related story: " + story + (s ? " · " + s.title : "") + "\n";
    }
    body += "Labels/tags: " + (tags || "none") + "\n";
    if (folder) body += "Suggested SharePoint folder: " + folder + "\n";
    if (desc) body += "\nDescription:\n" + desc + "\n";
    body += "\n----------------------------------------\n";
    body += file
      ? ("REMINDER: attach \"" + file.name + "\" to this email before sending — a webpage can't attach it automatically.\n")
      : "REMINDER: attach your document file(s) to this email before sending.\n";
    return { subject, body };
  }

  form.addEventListener("submit", (e) => e.preventDefault());

  if (sendBtn) {
    sendBtn.addEventListener("click", () => {
      if (!docNameInput.value.trim() && !fileInput.files.length) {
        alert("Please enter a document name or attach a file before sending.");
        return;
      }
      const { subject, body } = buildEmail();
      const mailto = "mailto:" + COE_INTAKE_EMAIL + "?subject=" + encodeURIComponent(subject) + "&body=" + encodeURIComponent(body);
      // Exposed for testing/inspection; the real action is the navigation below.
      window.__lastMailto = mailto;
      window.location.href = mailto;
    });
  }
}

function deleteUserDocument(id) {
  saveUserDocs(getUserDocs().filter(d => d.id !== id));
}

function updateDocTags(doc, newTags) {
  const today = new Date().toISOString().slice(0, 10);
  if (doc.sourceType === "seed") {
    const overrides = getSeedOverrides();
    overrides[doc.id] = { tags: newTags, lastModifiedBy: getUserName(), lastModifiedDate: today };
    saveSeedOverrides(overrides);
  } else {
    const docs = getUserDocs().map(d => d.id === doc.id ? { ...d, tags: newTags, lastModifiedBy: getUserName(), lastModifiedDate: today } : d);
    saveUserDocs(docs);
  }
  logActivity("tag", doc.name);
}

/**
 * "View in SharePoint" — always opens the real item/folder link. For
 * documents fetched live via Microsoft Graph, doc.url is the real item's
 * webUrl; otherwise it falls back to the shared folder link.
 */
function openDocument(doc) {
  bumpDownload(doc.id);
  logActivity("view", doc.name);
  window.open(doc.url || SHAREPOINT_FOLDER_URL, "_blank", "noopener");
}

/**
 * Download button.
 * - Graph-sourced documents (real SharePoint files, doc.sourceType==="graph"):
 *   fetches the real file bytes via Microsoft Graph and downloads them for
 *   real — see downloadGraphItem() in js/graph.js.
 * - Documents uploaded through this site's old local-demo storage (if any
 *   remain from before Graph was connected) still download from their saved
 *   base64 bytes.
 * - Anything else (seed docs before Graph was connected, or if the Graph
 *   download fails for any reason) falls back to opening SharePoint.
 */
async function downloadDocument(doc) {
  logActivity("download", doc.name);
  bumpDownload(doc.id);

  if (doc.sourceType === "graph" && typeof downloadGraphItem === "function") {
    try {
      await downloadGraphItem(doc);
      return;
    } catch (e) {
      console.error("Microsoft Graph download failed, falling back to SharePoint link:", e);
    }
  }

  if (doc.fileData) {
    const a = document.createElement("a");
    a.href = doc.fileData;
    a.download = doc.name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    return;
  }

  window.open(doc.url || SHAREPOINT_FOLDER_URL, "_blank", "noopener");
}

/* ---------------- story status (dashboard-only editing) ---------------- */
function getStatusOverrides() { return lsGet(LS.statusOverrides, {}); }
function saveStatusOverrides(obj) { lsSet(LS.statusOverrides, obj); }
function getStoriesWithStatus() {
  const overrides = getStatusOverrides();
  return STORIES.map(s => ({ ...s, status: overrides[s.code] || s.status }));
}
function setStoryStatus(code, status) {
  const overrides = getStatusOverrides();
  overrides[code] = status;
  saveStatusOverrides(overrides);
  logActivity("status", "Story " + code + " → " + status);
}

/* ---------------- unified search (stories + documents) ---------------- */
function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, s => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[s]));
}

function buildSnippet(text, words) {
  if (!text) return null;
  const lower = text.toLowerCase();
  let idx = -1;
  for (const w of words) {
    idx = lower.indexOf(w);
    if (idx !== -1) break;
  }
  if (idx === -1) return null;
  const start = Math.max(0, idx - 60);
  const end = Math.min(text.length, idx + 120);
  let snippet = text.slice(start, end).replace(/\s+/g, " ").trim();
  if (start > 0) snippet = "…" + snippet;
  if (end < text.length) snippet = snippet + "…";
  return snippet;
}

function searchAll(query, { pillar = "All", typeFilter = "All" } = {}) {
  const q = query.trim().toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  const pool = [...getStoriesWithStatus(), ...getAllDocuments()];

  let list = pool;
  if (pillar !== "All") list = list.filter(d => d.pillarCode === pillar);
  if (typeFilter !== "All") list = list.filter(d => d.type === typeFilter);

  if (!words.length) return list;

  const scored = list.map(item => {
    const metaText = [
      item.code || "", item.title || item.name || "", item.pillar, item.status || "",
      item.owner || item.uploadedBy || "", item.description || "", (item.tags || []).join(" ")
    ].join(" ").toLowerCase();
    const fullTextLower = (item.fullText || "").toLowerCase();

    let score = 0;
    let matchedInBody = false;
    for (const w of words) {
      if (metaText.includes(w)) {
        score += 1;
        if ((item.title || item.name || "").toLowerCase().includes(w)) score += 1;
        if ((item.code || "").toLowerCase() === w) score += 3;
      }
      // Full-text matches count too, but weighted lower than metadata —
      // finding a word once buried in a 10-page document is a weaker
      // signal than it appearing in the title or tags.
      if (fullTextLower && fullTextLower.includes(w)) {
        score += 0.5;
        matchedInBody = true;
      }
    }
    if (score <= 0) return null;
    const resultItem = matchedInBody ? { ...item, _snippet: buildSnippet(item.fullText, words) } : item;
    return { item: resultItem, score };
  }).filter(Boolean);

  scored.sort((a, b) => b.score - a.score);
  return scored.map(r => r.item);
}

document.addEventListener("DOMContentLoaded", () => {
  ensureSeedActivity();
  initRoleSwitcher();
  initNavDropdowns();
  initHeroSearch();
  renderQuickLinks();
  renderDocumentsPage();
  renderStoryDocSections();
  renderDashboardExtras();
  initRegisterForm();
});

/* ---------------- nav dropdown (mobile tap-to-open) ---------------- */
function initNavDropdowns() {
  document.querySelectorAll(".nav-item").forEach(item => {
    const trigger = item.querySelector(":scope > a");
    if (!trigger) return;
    trigger.addEventListener("click", (e) => {
      if (window.innerWidth > 900) return;
      e.preventDefault();
      item.classList.toggle("open");
    });
  });
  const toggle = document.querySelector(".nav-toggle");
  if (toggle) toggle.addEventListener("click", () => document.querySelector(".site-nav").classList.toggle("open"));
}

/* ---------------- home hero: search + pillar filter ---------------- */
let activePillar = "All";

function filterPillar(code, btn) {
  activePillar = code;
  document.querySelectorAll("#filterPills button").forEach(b => b.classList.remove("active"));
  if (btn) btn.classList.add("active");
  runHeroSearch();
}

function runHeroSearch() {
  const box = document.getElementById("heroResults");
  const input = document.getElementById("searchInput");
  if (!box || !input) return;
  const q = input.value.trim();

  if (!q && activePillar === "All") {
    box.classList.remove("show");
    box.innerHTML = "";
    return;
  }

  const list = searchAll(q, { pillar: activePillar });
  box.classList.add("show");
  if (list.length === 0) {
    box.innerHTML = `<div class="results-empty">No results. Try a different pillar or keyword.</div>`;
    return;
  }
  box.innerHTML = list.slice(0, 14).map(renderResultCard).join("");
}

function renderResultCard(item) {
  const isStory = item.type === "story";
  const title = isStory ? item.title : item.name;
  const meta = isStory ? `${escapeHtml(item.pillar)} · Owner: ${escapeHtml(item.owner)}` : `${escapeHtml(item.pillar || "Unlinked")} · ${escapeHtml(item.location || "SharePoint")}`;
  const desc = isStory ? item.description : `Tags: ${(item.tags || []).join(", ") || "—"}`;
  const link = isStory ? item.url : "#";
  const onclick = isStory ? "" : `onclick="openDocument(${JSON.stringify(item).replace(/"/g, '&quot;')});return false;"`;
  const statusChip = isStory ? `<span class="chip ${item.status.toLowerCase().replace(/\s+/g, '')}">${escapeHtml(item.status)}</span>` : "";
  const snippetHtml = item._snippet
    ? `<p class="body-match">🔍 Matched inside the document: “${escapeHtml(item._snippet)}”</p>`
    : "";
  return `
    <div class="result-card">
      <div class="rc-top">
        <a class="rc-title" href="${link}" ${onclick}>${item.code ? escapeHtml(item.code) + " · " : ""}${escapeHtml(title)}</a>
        <span style="display:flex;gap:6px;align-items:center">
          <span class="type-badge ${item.type}">${item.type}</span>${statusChip}
        </span>
      </div>
      <div class="rc-meta">${meta}</div>
      <p>${escapeHtml(desc)}</p>
      ${snippetHtml}
    </div>`;
}

function initHeroSearch() {
  const input = document.getElementById("searchInput");
  if (!input) return;
  input.addEventListener("input", runHeroSearch);
}

/* ---------------- home: quick links ---------------- */
function renderQuickLinks() {
  const root = document.getElementById("quickLinks");
  if (!root) return;
  const all = getAllDocuments();

  const featured = all.filter(d => d.featured).slice(0, 4);
  const mostDownloaded = [...all].sort((a, b) => b.downloads - a.downloads).slice(0, 4);
  const recentlyModified = [...all].sort((a, b) => (b.lastModifiedDate || "").localeCompare(a.lastModifiedDate || "")).slice(0, 4);

  const col = (title, items, sub) => `
    <div class="ql-card">
      <h4>${title}</h4>
      ${items.length ? `<ul>${items.map(d => `<li><a href="#" onclick='openDocument(${JSON.stringify(d).replace(/'/g, "&apos;")});return false;'>${escapeHtml(d.name)}</a><span>${sub(d)}</span></li>`).join("")}</ul>` : `<div class="ql-empty">Nothing here yet.</div>`}
    </div>`;

  root.innerHTML = [
    col("Featured Documents", featured, d => d.pillar || "General"),
    col("Most Downloaded", mostDownloaded, d => d.downloads + " opens"),
    col("Recently Modified", recentlyModified, d => "Updated " + d.lastModifiedDate)
  ].join("");
}

/* ---------------- documents.html: full catalog ---------------- */
function renderDocumentsPage() {
  const root = document.getElementById("docCatalog");
  if (!root) return;

  const searchBox = document.getElementById("docSearch");
  const pillarSelect = document.getElementById("docPillarFilter");

  function draw() {
    const q = searchBox ? searchBox.value : "";
    const pillar = pillarSelect ? pillarSelect.value : "All";
    let list = q ? searchAll(q, { pillar, typeFilter: "document" }) : getAllDocuments().filter(d => pillar === "All" || d.pillarCode === pillar);
    list.sort((a, b) => (b.lastModifiedDate || "").localeCompare(a.lastModifiedDate || ""));
    root.innerHTML = list.length ? list.map(d => docRowHtml(d)).join("") : `<div class="results-empty">No documents match.</div>`;
  }

  if (searchBox) searchBox.addEventListener("input", draw);
  if (pillarSelect) pillarSelect.addEventListener("change", draw);
  draw();
}

function docRowHtml(d) {
  const canManage = d.sourceType === "user";
  const isPending = d.sourceType === "pending";
  const canRealDownload = !!d.fileData || d.sourceType === "graph";
  let indexBadge = "";
  if (d.fullText) {
    indexBadge = '<span class="type-badge story" title="Every line of this document is searchable">🔍 full-text indexed</span>';
  } else if ((d.sourceType === "user" || isPending) && d.fullTextStatus === "unsupported") {
    indexBadge = '<span class="sso-note">(full-text search not available for this file type)</span>';
  }
  const sharePointOnlyTag = (canRealDownload || isPending) ? "" : '<span class="sso-note">(SharePoint only)</span>';
  const pendingTag = isPending ? '<span class="type-badge document" title="Attached on the Register a Document page — not yet emailed">📋 pending submission</span>' : "";
  const actions = isPending
    ? `<span class="sso-note">Fill out and send the form above to submit this ↑</span>`
    : `<button onclick='downloadDocument(${JSON.stringify(d).replace(/'/g, "&apos;")})'>⬇ Download</button>
       <button onclick='openDocument(${JSON.stringify(d).replace(/'/g, "&apos;")})'>↗ View in SharePoint</button>
       ${canManage ? `<button class="danger contributor-only" onclick="handleDeleteDoc('${d.id}')">Remove</button>` : ""}`;
  return `
    <div class="doc-row" data-doc-id="${d.id}">
      <div class="dr-main">
        <div class="dr-name">📄 ${escapeHtml(d.name)} ${d.featured ? '<span class="type-badge document">featured</span>' : ""}${sharePointOnlyTag}${pendingTag} ${indexBadge}</div>
        <div class="dr-meta">${escapeHtml(d.pillar || "Unlinked")} · Uploaded by ${escapeHtml(d.uploadedBy)} on ${d.uploadDate} · Last modified by ${escapeHtml(d.lastModifiedBy)} on ${d.lastModifiedDate} · ${d.downloads} opens</div>
        <div class="dr-tags" id="tags-${d.id}">${renderTagPills(d)}</div>
        ${d._snippet ? `<p class="body-match" style="margin-top:8px">🔍 Matched inside the document: “${escapeHtml(d._snippet)}”</p>` : ""}
        <div class="contributor-only tag-add-form">
          <input type="text" placeholder="add tag…" id="newtag-${d.id}">
          <button type="button" onclick="handleAddTag('${d.id}')">Add tag</button>
        </div>
      </div>
      <div class="dr-actions">
        ${actions}
      </div>
    </div>`;
}

function renderTagPills(d) {
  return (d.tags || []).map(t => `<span class="tag-pill">${escapeHtml(t)}<span class="rm contributor-only" onclick="handleRemoveTag('${d.id}','${escapeHtml(t)}')">&times;</span></span>`).join("") || `<span class="ql-empty">No tags</span>`;
}

function findDocById(id) { return getAllDocuments().find(d => d.id === id); }

function handleAddTag(id) {
  const input = document.getElementById("newtag-" + id);
  const val = input.value.trim();
  if (!val) return;
  const doc = findDocById(id);
  if (!doc) return;
  const tags = Array.from(new Set([...(doc.tags || []), val]));
  updateDocTags(doc, tags);
  input.value = "";
  refreshDocViews();
}
function handleRemoveTag(id, tag) {
  const doc = findDocById(id);
  if (!doc) return;
  updateDocTags(doc, (doc.tags || []).filter(t => t !== tag));
  refreshDocViews();
}
function handleDeleteDoc(id) {
  if (!confirm("Remove this document entry? (This only removes the catalog entry from your browser — nothing is deleted in SharePoint.)")) return;
  deleteUserDocument(id);
  refreshDocViews();
}
function refreshDocViews() {
  renderDocumentsPage();
  renderStoryDocSections();
  renderQuickLinks();
}

/* ---------------- pillar pages: per-story document sections ---------------- */
function renderStoryDocSections() {
  document.querySelectorAll("[data-story-docs]").forEach(container => {
    const code = container.getAttribute("data-story-docs");
    const story = STORIES.find(s => s.code === code);
    const docs = getDocsForStory(code);
    const registerHref = `register-document.html?pillar=${story ? story.pillarCode : ""}&story=${code}`;
    container.innerHTML = `
      <h4>Documents for this deliverable</h4>
      ${docs.length ? docs.map(docRowHtml).join("") : `<div class="ql-empty">No documents linked yet.</div>`}
      <a class="doc-link" href="${registerHref}" style="margin-top:10px">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
        Register a document for this deliverable
      </a>`;
  });
}

/* ---------------- dashboard: metrics, activity tracker, story table ---------------- */
function renderDashboardExtras() {
  renderMetrics();
  renderActivityTracker();
  renderStoryTable();
}

function renderMetrics() {
  const root = document.getElementById("dashMetrics");
  if (!root) return;
  const docs = getAllDocuments();
  const totalDownloads = docs.reduce((sum, d) => sum + (d.downloads || 0), 0);
  const activity = getActivity();
  const contributors = new Set(activity.filter(a => ["upload", "tag", "status"].includes(a.type)).map(a => a.actor));
  root.innerHTML = `
    <div class="metric-card"><div class="n">${docs.length}</div><div class="l">Documents in repository</div></div>
    <div class="metric-card"><div class="n">${totalDownloads}</div><div class="l">Total opens / downloads</div></div>
    <div class="metric-card"><div class="n">${contributors.size}</div><div class="l">Active contributors</div></div>`;
}

function renderActivityTracker() {
  const barsRoot = document.getElementById("activityBars");
  const logRoot = document.getElementById("activityLog");
  if (!barsRoot && !logRoot) return;
  const activity = getActivity();

  const days = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    days.push(d.toISOString().slice(0, 10));
  }
  const counts = days.map(day => activity.filter(a => a.timestamp === day).length);
  const max = Math.max(1, ...counts);

  if (barsRoot) {
    barsRoot.innerHTML = days.map((day, i) => {
      const h = Math.round((counts[i] / max) * 90) + 4;
      const label = new Date(day).toLocaleDateString(undefined, { weekday: "short" });
      return `<div class="ab-col"><div class="ab-count">${counts[i]}</div><div class="ab-bar" style="height:${h}px"></div><div class="ab-label">${label}</div></div>`;
    }).join("");
  }

  if (logRoot) {
    const recent = [...activity].reverse().slice(0, 12);
    const verbs = { upload: "uploaded", download: "downloaded", view: "viewed", tag: "updated tags on", status: "updated" };
    logRoot.innerHTML = recent.map(a => `<div class="al-item"><span>${escapeHtml(a.actor)} ${verbs[a.type] || a.type} <b>${escapeHtml(a.target)}</b></span><span>${a.timestamp}</span></div>`).join("") || `<div class="ql-empty">No activity yet.</div>`;
  }
}

function renderStoryTable() {
  const root = document.getElementById("storyTableBody");
  if (!root) return;
  const stories = getStoriesWithStatus();
  const statuses = ["Done", "In Progress", "To Do", "Backlog"];
  root.innerHTML = stories.map(s => `
    <tr>
      <td><a href="${s.url}">${s.code}</a></td>
      <td>${escapeHtml(s.title)}</td>
      <td>${escapeHtml(s.pillar)}</td>
      <td>${escapeHtml(s.owner)}</td>
      <td>
        <span class="viewer-only-status chip ${s.status.toLowerCase().replace(/\s+/g, '')}">${escapeHtml(s.status)}</span>
        <select class="contributor-only" style="display:none" onchange="handleStatusChange('${s.code}', this.value)">
          ${statuses.map(st => `<option value="${st}" ${st === s.status ? "selected" : ""}>${st}</option>`).join("")}
        </select>
      </td>
    </tr>`).join("");
  syncStatusEditors();
}

function syncStatusEditors() {
  const isContributor = getRole() === "contributor";
  document.querySelectorAll(".story-table select.contributor-only").forEach(s => s.style.display = isContributor ? "inline-block" : "none");
  document.querySelectorAll(".story-table .viewer-only-status").forEach(s => s.style.display = isContributor ? "none" : "inline-block");
}

function handleStatusChange(code, status) {
  setStoryStatus(code, status);
  renderStoryTable();
  renderActivityTracker();
}

/* re-sync role-dependent UI whenever role changes */
/* role changes automatically re-sync status editors via setRole() above */

/* ---------------- chatbot ---------------- */
let chatbotOpen = false;
function toggleChatbot() {
  const win = document.getElementById("chatbot-window");
  chatbotOpen = !chatbotOpen;
  win.classList.toggle("open", chatbotOpen);
}
function sendMessage() {
  const input = document.getElementById("chat-input");
  const message = input.value.trim();
  if (message === "") return;
  addUserMessage(message);
  answerFromKnowledgeBase(message);
  input.value = "";
}
function addUserMessage(message) {
  const body = document.getElementById("chat-body");
  body.innerHTML += `<div class="user-message">${escapeHtml(message)}</div>`;
  body.scrollTop = body.scrollHeight;
}
function addBotMessage(html) {
  const body = document.getElementById("chat-body");
  body.innerHTML += `<div class="bot-message">${html}</div>`;
  body.scrollTop = body.scrollHeight;
}
function answerFromKnowledgeBase(keyword) {
  const results = searchAll(keyword);
  if (results.length === 0) {
    addBotMessage("I couldn't find a match. Try a pillar name, a deliverable code (e.g. \"3.6\"), a document keyword, or even a phrase from inside an uploaded document.");
    return;
  }
  let html = "<b>I found these:</b><ul>";
  results.slice(0, 6).forEach(item => {
    const isStory = item.type === "story";
    const title = isStory ? `${item.code} ${item.title}` : item.name;
    const href = isStory ? item.url : (item.url || SHAREPOINT_FOLDER_URL);
    html += `<li>${isStory ? "📁" : "📄"} <a href="${href}" target="${isStory ? "_self" : "_blank"}">${escapeHtml(title)}</a> <span style="color:#5B6B7A">(${item.type})</span>`;
    if (item._snippet) {
      html += `<br><span style="color:#5B6B7A;font-size:11.5px">🔍 “${escapeHtml(item._snippet)}”</span>`;
    }
    html += `</li>`;
  });
  html += "</ul>";
  addBotMessage(html);
}

/* ---------------- about modal ---------------- */
function toggleAbout() {
  const overlay = document.getElementById("about-overlay");
  const modal = document.getElementById("about-modal");
  if (!overlay || !modal) return;
  const isOpen = modal.classList.contains("open");
  overlay.classList.toggle("open", !isOpen);
  modal.classList.toggle("open", !isOpen);
}
