# D S Seeds Website

Corporate / informational website for **D S Seeds (OPC) Private Limited**, an agricultural seeds company, built by Embetronix Technologies.

Start here: `CONTEXT.md` (project facts and working agreements) and `HANDOVER.md` (current status and next steps).

## Regenerating documents
```
npm install docx
node scripts/build_questionnaire.js      # English
node scripts/build_questionnaire_bn.js   # Bangla
```
Output: `docs/01-requirements/DS_Seeds_Website_Requirements_Questionnaire_v<VERSION>[_BN].docx`
