const templates = {
  welcome: {
    text: "Hi {{name}},\nWelcome to {{company}}!",
    variables: ["name", "company"]
  },
  followup: {
    text: "Hello {{client}},\nJust following up on {{subject}}.",
    variables: ["client", "subject"]
  }
};

document.addEventListener("DOMContentLoaded", () => {
  const list = document.getElementById("templateList");
  const form = document.getElementById("variableForm");
  const insertBtn = document.getElementById("insertBtn");

  for (const key in templates) {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = key;
    list.appendChild(opt);
  }

  list.addEventListener("change", () => {
    form.innerHTML = "";
    const vars = templates[list.value].variables;
    vars.forEach(v => {
      const input = document.createElement("input");
      input.placeholder = v;
      input.name = v;
      input.style.display = "block";
      form.appendChild(input);
    });
  });

  insertBtn.addEventListener("click", () => {
    const selected = templates[list.value];
    let filled = selected.text;
    selected.variables.forEach(v => {
      const val = document.querySelector(`[name=${v}]`).value;
      filled = filled.replace(`{{${v}}}`, val);
    });

    Office.context.mailbox.item.body.setSelectedDataAsync(filled, { coercionType: Office.CoercionType.Text });
  });

  list.dispatchEvent(new Event("change")); // load default
});
