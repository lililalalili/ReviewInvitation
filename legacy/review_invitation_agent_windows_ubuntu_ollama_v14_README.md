# v14 commercialization-oriented upgrade

This version is generated from v13 and upgrades four high-priority areas:

1. Unified network layer
   - shared requests Session
   - retry / backoff / timeout handling
   - used for PubMed and Ollama/Kimi calls

2. Better email extraction
   - scans all affiliations
   - prefers affiliations with 'Electronic address:'
   - falls back to any affiliation containing an email

3. PMID written into main sheet
   - new column 25 = PMID
   - improves traceability and future dedup/update workflows

4. SQLite state database
   - replaces processed_pmids.json
   - state file: review_invitation_agent_state.sqlite3
   - improves reliability over a flat JSON file

Also included:
- structured neuro decision source directly from decision logic
- stronger duplicate checking using existing PMID values in the sheet
- existing v13 run lock
- existing retry for log append
