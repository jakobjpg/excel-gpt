# Excel GPT - Project Requirements

## Overview
Excel GPT is an Excel add-in that enables advanced cell manipulation and automation using JSON and agent-based logic, powered by OpenAI.

---

## Requirements Checklist

- [x] **Playground Area Selection**
  - User can select a range in Excel to define the 'playground' area where all operations are allowed.

- [ ] **Get/Set Cells via JSON**
  - [ ] Set format of a cell (font, color, etc.)
  - [ ] Set content of a cell (formulas, raw numbers, etc.)
  - [ ] Get content from individual cell (e.g., B2) including formulas, raw numbers, etc.
  - [ ] Get all content from cells in the playground area

- [ ] **Agent Functionality**
  - [ ] Agent can understand a task, plan out the task, and execute it (in any order)
  - [ ] Agent can perform a batch of changes using a JSON array
  - [ ] Agent can perform a single change at a time

- [x] **React + Office.js Integration**
  - Project is scaffolded with React and Office.js, and playground selection is implemented in React.

- [x] **Development Environment**
  - Project builds and runs with React 18+ and modern Babel/Webpack config.

---

## Notes
- All get/set operations must be restricted to the selected playground area.
- The agent will use the OpenAI API to interpret and plan tasks.
- No UI is required for get/set functions; they are callable by the agent logic.

---

## Next Steps
- Implement JSON-based get/set cell functions (content and formatting)
- Implement agent logic for planning and executing tasks
- Integrate OpenAI API for agent intelligence 