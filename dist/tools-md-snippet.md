<!-- m365-agent-cli:tools-md begin -->
## m365-agent-cli

Microsoft 365 CLI for calendar, mail, OneDrive, Planner, SharePoint, To Do, Teams, Bookings, and more.

**Installation:**
```bash
npm install m365-agent-cli
```

**OpenClaw Skill:** After installing via npm, the skill is available at:
```
node_modules/m365-agent-cli/skills/m365-agent-cli/SKILL.md
```

To install the skill for OpenClaw:
```bash
mkdir -p ~/.openclaw/workspace/skills
cp -r node_modules/m365-agent-cli/skills/m365-agent-cli ~/.openclaw/workspace/skills/
```

Or set `OPENCLAW_SKILLS_DIR` environment variable and run the postinstall hook (see package documentation).

**Documentation:** https://github.com/markus-lassfolk/m365-agent-cli

**Version:** Matches the installed npm package version (run `npm list m365-agent-cli` to check).
<!-- m365-agent-cli:tools-md end -->
