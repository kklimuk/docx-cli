You are testing **docx-cli**, a command-line tool for reading and editing Microsoft Word (.docx) files. You have NOT used it before — discover what you need from its own `--help`; do not assume flags.

You have a **bash** tool. To read a text file, run `cat <path>`. To run docx-cli, invoke the executable at this exact path:
  {{BINARY}}

First, orient yourself by running these with bash:
  {{BINARY}} --help
  {{BINARY}} info locators
  {{BINARY}} <command> --help     (for any command you decide to use)

Also list your assets folder (it may be empty); if it holds files, read them:
  ls assets

{{WORKLINE}}
  {{DOC}}

## Your task

{{TASK}}

## Your job
Carry out the task above on the working document. The request describes the OUTCOME the person wants — YOU work out which docx-cli commands get there.

## Rules
- Use ONLY the docx-cli executable above for document operations. Do NOT unzip the .docx, hand-edit XML, or use any other tool for the document.
- STAY IN YOUR CURRENT FOLDER. Do not read or list files outside it — no `ls -R /`, no exploring parent directories. Everything you need is here.
- Locators like p0, t0:r1c2:p0, sN shift after structural edits — re-read the document (`{{BINARY}} read {{DOC}}`) when unsure. Prefer batch commands where the tool offers them.
- If a command fails, read its error, try a couple of sensible fixes, then move on. Do not loop on one step forever.
- Finish the task if you can.

When you are done, write ONE short paragraph summarizing what you changed.
