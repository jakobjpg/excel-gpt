import React, { useState } from "react";

// Utility: Check if a cell address is within the playground range
function isCellInRange(cellAddress, playgroundRange) {
  // Basic implementation: check if address is within the playground range
  // This only works for simple ranges like 'A1:D10' and addresses like 'B2'
  if (!playgroundRange) return false;
  const match = playgroundRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (!match) return false;
  const [, startCol, startRow, endCol, endRow] = match;
  const colToNum = col => col.split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  const addrMatch = cellAddress.match(/([A-Z]+)(\d+)/);
  if (!addrMatch) return false;
  const [, addrCol, addrRow] = addrMatch;
  const colNum = colToNum(addrCol);
  const startColNum = colToNum(startCol);
  const endColNum = colToNum(endCol);
  const rowNum = parseInt(addrRow, 10);
  const startRowNum = parseInt(startRow, 10);
  const endRowNum = parseInt(endRow, 10);
  return (
    colNum >= Math.min(startColNum, endColNum) &&
    colNum <= Math.max(startColNum, endColNum) &&
    rowNum >= Math.min(startRowNum, endRowNum) &&
    rowNum <= Math.max(startRowNum, endRowNum)
  );
}

// Get content and format from a cell (address as 'A1', 'B2', etc.)
export async function getCell(address, playgroundRange) {
  if (!isCellInRange(address, playgroundRange)) throw new Error("Cell outside playground area");
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.load(["values", "formulas", "format/font/color", "format/fill/color"]);
    await context.sync();
    return {
      address,
      value: range.values[0][0],
      formula: range.formulas[0][0],
      fontColor: range.format.font.color,
      fillColor: range.format.fill.color,
    };
  });
}

// Set content/format of a cell
export async function setCell(address, data, playgroundRange) {
  if (!isCellInRange(address, playgroundRange)) throw new Error("Cell outside playground area");
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    if (data.value !== undefined) range.values = [[data.value]];
    if (data.formula !== undefined) range.formulas = [[data.formula]];
    // Support bold formatting
    if (data.bold !== undefined) range.format.font.bold = !!data.bold;
    // Support font color
    if (data.fontColor && data.fontColor !== 'bold') range.format.font.color = data.fontColor;
    // Support fill color
    if (data.fillColor) range.format.fill.color = data.fillColor;
    // Support font name
    if (data.fontName) range.format.font.name = data.fontName;
    // Support font size
    if (data.fontSize) range.format.font.size = data.fontSize;
    // If fontColor is 'bold', treat as bold
    if (data.fontColor === 'bold') range.format.font.bold = true;
    await context.sync();
    return true;
  });
}

// Get all content in the playground area
export async function getAllContent(playgroundRange) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(playgroundRange);
    range.load(["address", "values", "formulas", "format/font/color", "format/fill/color"]);
    await context.sync();
    // Flatten to array of cell objects
    const result = [];
    const rows = range.values.length;
    const cols = range.values[0].length;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        result.push({
          address: `${String.fromCharCode(65 + c)}${r + 1}`,
          value: range.values[r][c],
          formula: range.formulas[r][c],
          fontColor: range.format.font.color,
          fillColor: range.format.fill.color,
        });
      }
    }
    return result;
  });
}

// Agent: Understands a task, plans, and executes using the above functions
async function agentExecuteTask(task, playgroundRange) {
  // Example: task = { actions: [ { type: 'set', address: 'B2', data: { value: 42 } }, ... ] }
  for (const action of task.actions) {
    if (action.type === "set") {
      await setCell(action.address, action.data, playgroundRange);
    } else if (action.type === "get") {
      await getCell(action.address, playgroundRange);
    }
    // Add more action types as needed
  }
  return true;
}

// Helper to map alternative action formats to the expected format
function normalizeAction(action) {
  // If already in the expected format
  if (action.type && action.address && action.data) return action;
  // Map from { action: 'setCellValue', cell: 'B1', value: 'Country' }
  if (action.action === 'setCellValue' && action.cell && action.value !== undefined) {
    return { type: 'set', address: action.cell, data: { value: action.value } };
  }
  // Map from { action: 'setCellFormula', cell: 'B1', formula: '...' }
  if (action.action === 'setCellFormula' && action.cell && action.formula !== undefined) {
    return { type: 'set', address: action.cell, data: { formula: action.formula } };
  }
  // Add more mappings as needed
  return action;
}

// Helper to build context from chat history
function buildContextFromChat(chat) {
  // Only include user and agent messages for context
  return chat.filter(msg => msg.sender === "user" || msg.sender === "agent").map(msg => ({ sender: msg.sender, text: msg.text }));
}

// Two-step agent: reasoning, then JSON-only agent
async function agentProcessMessage(message, playgroundRange, setChat, setActionsJson, setShowActions, context = [], depth = 0, chatHistory = null) {
  const MAX_DEPTH = 10;
  if (depth > MAX_DEPTH) {
    setChat((prev) => [...prev, { sender: "system", text: `Agent reached maximum recursion depth (${MAX_DEPTH}).` }]);
    return;
  }
  // Use full chat history for context if available
  let fullContext = context;
  if (chatHistory) {
    fullContext = buildContextFromChat(chatHistory);
  }
  if (fullContext.length > 0 && fullContext[fullContext.length - 1]?.text === message) {
    setChat((prev) => [...prev, { sender: "system", text: "Agent repeated itself. Stopping recursion." }]);
    return;
  }
  setChat((prev) => [...prev, { sender: depth === 0 ? "user" : "agent", text: message }]);

  // Always fetch the latest playground cell data before each step
  let playgroundSummary = "";
  try {
    const allCells = await getAllContent(playgroundRange);
    if (allCells && allCells.length > 0) {
      playgroundSummary =
        "Current playground area cell data (address, value, formula, fontColor, fillColor):\n" +
        allCells.map(cell => `${cell.address}: value=${cell.value}, formula=${cell.formula}`).join("\n");
    } else {
      playgroundSummary = "Playground area is empty.";
    }
  } catch (e) {
    playgroundSummary = "[Error reading playground area]";
  }

  // Step 1: Reasoning agent
  const conventions = `
Conventions:
- Each cell should contain only one type of data (number or string, not both).
- For example, there should never be a cell with 'United States - GDP 21.43 trillion'.
- Instead, use one cell for the country name (e.g., 'United States'), and another cell for the GDP value (e.g., '21.43'), with column headers like 'Country' and 'GDP (trillion)'.
- The first row should contain headers for each column, and headers should be bolded.`;
  const reasoningPrompt = `You are Excel GPT. The user has selected the playground area: ${playgroundRange}.
${playgroundSummary}
${conventions}
Here is the conversation so far:
${fullContext.map((c, i) => `${c.sender}: ${c.text}`).join("\n")}
Task: ${message}
Please reason step by step and describe your plan for how to accomplish the task. Do not output any JSON yet. Never ask the user for confirmation or permission; always proceed with the actions you suggest.`;
  try {
    // Step 1: Get reasoning and plan
    const reasoningResponse = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${process.env.OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: "gpt-3.5-turbo",
        messages: [{ role: "user", content: reasoningPrompt }],
        temperature: 0,
      }),
    });
    const reasoningData = await reasoningResponse.json();
    const reasoningContent = reasoningData.choices?.[0]?.message?.content || "[Error: No response from OpenAI]";
    setChat((prev) => [...prev, { sender: "agent", text: reasoningContent }]);

    // Always fetch the latest playground cell data before JSON step
    let playgroundSummary2 = "";
    try {
      const allCells2 = await getAllContent(playgroundRange);
      if (allCells2 && allCells2.length > 0) {
        playgroundSummary2 =
          "Current playground area cell data (address, value, formula, fontColor, fillColor):\n" +
          allCells2.map(cell => `${cell.address}: value=${cell.value}, formula=${cell.formula}`).join("\n");
      } else {
        playgroundSummary2 = "Playground area is empty.";
      }
    } catch (e) {
      playgroundSummary2 = "[Error reading playground area]";
    }

    // Step 2: JSON-only agent
    const jsonPrompt = `You are an Excel automation agent. Based on the following plan and the current state of the playground area, output only a JSON array of actions to perform, and nothing else. Use the format: [{\"type\":\"set\",\"address\":\"B2\",\"data\":{\"value\":42}}]. To set a cell as bold, use { bold: true } in the data object. To set font color, use { fontColor: '#RRGGBB' }. To set fill color, use { fillColor: '#RRGGBB' }. To set font name or size, use { fontName: 'Arial', fontSize: 12 }. If you want to set a formula, use data: { formula: ... }. Here is the plan and context:\n${reasoningContent}\n\n${playgroundSummary2}\nNever ask the user for confirmation or permission; always proceed with the actions you suggest.`;
    const jsonResponse = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${process.env.OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: "gpt-3.5-turbo",
        messages: [{ role: "user", content: jsonPrompt }],
        temperature: 0,
      }),
    });
    const jsonData = await jsonResponse.json();
    const jsonContent = jsonData.choices?.[0]?.message?.content || "[Error: No response from OpenAI]";
    // Try to extract the first JSON array from the response
    const jsonArrayMatch = jsonContent.match(/\[\s*{[\s\S]*?}\s*\]/);
    let actions;
    if (jsonArrayMatch) {
      try {
        actions = JSON.parse(jsonArrayMatch[0]);
        // Normalize actions
        actions = actions.map(normalizeAction);
      } catch {
        setChat((prev) => [...prev, { sender: "system", text: "Could not parse agent actions as JSON." }]);
        return;
      }
      // Step 3: Execute actions
      try {
        for (const action of actions) {
          if (!isCellInRange(action.address, playgroundRange)) {
            setChat((prev) => [...prev, { sender: "system", text: `Skipped action for ${action.address}: outside playground area.` }]);
            continue;
          }
          if (action.type === "set") {
            await setCell(action.address, action.data, playgroundRange);
          } else if (action.type === "get") {
            const cell = await getCell(action.address, playgroundRange);
            setChat((prev) => [...prev, { sender: "system", text: `Cell ${action.address}: ${JSON.stringify(cell)}` }]);
          } else if (action.type === "clear") {
            await setCell(action.address, { value: "" }, playgroundRange);
          }
          // Add more action types as needed
        }
        setActionsJson(actions);
        setShowActions(false);
        setChat((prev) => [...prev, { sender: "system", text: `Executed ${actions.length} actions. (Click to view JSON)`, isActions: true }]);
        // Confirmation step: read playground area and summarize
        try {
          const allCellsAfter = await getAllContent(playgroundRange);
          const summary = allCellsAfter && allCellsAfter.length > 0
            ? allCellsAfter.map(cell => `${cell.address}: value=${cell.value}, formula=${cell.formula}`).join("\n")
            : "Playground area is empty.";
          setChat((prev) => [...prev, { sender: "agent", text: `Task complete: Here is the current state of the playground area after execution:\n${summary}` }]);
        } catch (e) {
          setChat((prev) => [...prev, { sender: "agent", text: `Task complete, but failed to read playground area for confirmation.` }]);
        }
      } catch (err) {
        setActionsJson(actions);
        setShowActions(false);
        setChat((prev) => [
          ...prev,
          { sender: "system", text: `Error: ${err.message} (Click to view attempted JSON)`, isActions: true }
        ]);
        return;
      }
    } else {
      // No JSON array found, treat as reasoning step and recurse
      setChat((prev) => [...prev, { sender: "agent", text: jsonContent }]);
      await agentProcessMessage(message, playgroundRange, setChat, setActionsJson, setShowActions, [...fullContext, { sender: "agent", text: jsonContent }], depth + 1, chatHistory);
    }
  } catch (err) {
    setChat((prev) => [...prev, { sender: "system", text: `Error: ${err.message}` }]);
  }
}

const App = () => {
  const [playgroundRange, setPlaygroundRange] = useState("");
  const [selecting, setSelecting] = useState(false);
  const [chat, setChat] = useState([]);
  const [input, setInput] = useState("");
  const [sending, setSending] = useState(false);
  const [actionsJson, setActionsJson] = useState(null);
  const [showActions, setShowActions] = useState(false);

  // Clear chat handler
  const handleClearChat = () => {
    setChat([]);
    setActionsJson(null);
    setShowActions(false);
  };

  const handleSelectPlayground = async () => {
    setSelecting(true);
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        setPlaygroundRange(range.address);
      });
    } catch (error) {
      console.error(error);
      setPlaygroundRange("");
    }
    setSelecting(false);
  };

  const handleSend = async (e) => {
    e.preventDefault();
    if (!input.trim() || !playgroundRange) return;
    setSending(true);
    await agentProcessMessage(input, playgroundRange, setChat, setActionsJson, setShowActions, [], 0, chat);
    setInput("");
    setSending(false);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", width: "100%" }}>
      <h2>Excel GPT</h2>
      <div style={{ margin: "16px 0", width: "100%" }}>
        <button onClick={handleSelectPlayground} disabled={selecting}>
          {selecting ? "Selecting..." : "Select Playground Area"}
        </button>
        {playgroundRange && (
          <div style={{ marginTop: 8 }}>
            <strong>Playground Area:</strong> {playgroundRange}
          </div>
        )}
        <button onClick={handleClearChat} style={{ float: "right", marginLeft: 8 }}>
          Clear Chat
        </button>
      </div>
      <div style={{ flex: 'none', height: 300, border: "1px solid #ccc", borderRadius: 4, padding: 8, overflowY: "auto", background: "#fafafa", width: "100%", minHeight: 0 }}>
        {chat.map((msg, i) => (
          <div
            key={i}
            style={{ margin: "4px 0", color: msg.sender === "user" ? "#0078d4" : msg.sender === "agent" ? "#333" : "#b00", cursor: msg.isActions ? "pointer" : "default", width: "100%" }}
            onClick={() => {
              if (msg.isActions && actionsJson) setShowActions((prev) => !prev);
            }}
          >
            <strong>{msg.sender}:</strong> {msg.text}
            {msg.isActions && showActions && actionsJson && (
              <pre style={{ background: "#eee", padding: 8, borderRadius: 4, marginTop: 4, whiteSpace: "pre-wrap", width: "100%" }}>
                {JSON.stringify(actionsJson, null, 2)}
              </pre>
            )}
          </div>
        ))}
      </div>
      <form onSubmit={handleSend} style={{ marginTop: 8, display: "flex", gap: 4, width: "100%" }}>
        <input
          type="text"
          value={input}
          onChange={e => setInput(e.target.value)}
          placeholder="Type your task for Excel GPT..."
          style={{ flex: 1 }}
          disabled={sending}
        />
        <button type="submit" disabled={sending || !input.trim() || !playgroundRange}>
          {sending ? "Sending..." : "Send"}
        </button>
      </form>
    </div>
  );
};

export default App; 