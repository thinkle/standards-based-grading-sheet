description: "Agent works through a TODO list until all items are complete or blocked."
tools: [codebase, terminal] # whatever you want available
constraints:

- Stick only to tasks in TODO.md
- Do not invent new tasks unless subtasking
- After each task, update status and move to next
- DO NOT COMMIT or push changes -- version control is how a human reviews your work and decides what to accept.
- ONLY edit files locally in the working directory.
  loop:
- read TODO.md
- while (unfinished tasks):
  - select next task
  - execute with available tools
  - update task status
  - log summary
