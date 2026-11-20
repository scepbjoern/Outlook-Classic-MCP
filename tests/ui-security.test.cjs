// Exercise the shipped inline scripts with an in-memory parent and DOM sinks.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

function view(name) {
  const elements = new Map();
  const sent = [];
  const parent = { postMessage(message, target) { sent.push({message, target}); } };
  let receive;
  function element(id) {
    if (!elements.has(id)) elements.set(id, {
      innerHTML: "", textContent: "", listeners: {},
      addEventListener(event, handler) { this.listeners[event] = handler; },
    });
    return elements.get(id);
  }
  const window = {parent, addEventListener(_, handler) { receive = handler; }};
  const context = vm.createContext({
    window, document: {getElementById: element, documentElement: {dataset: {}}, body: {}},
    ResizeObserver: class { observe() {} }, requestAnimationFrame() { return 1; },
  });
  const resource = file => fs.readFileSync(path.join(__dirname, "../src/outlook_mcp/ui", file), "utf8");
  const html = resource(name + ".html").replace("<!--%COMMON%-->", () => resource("_common.html"));
  for (const script of html.matchAll(/<script>([\s\S]*?)<\/script>/g)) {
    vm.runInContext(script[1], context, {filename: name});
  }
  function deliver(data, source = parent, origin = "https://host.example.test") {
    receive({source, origin, data: {jsonrpc: "2.0", ...data}});
  }
  function result(data, ...sender) {
    deliver({method: "ui/notifications/tool-result", params: {structuredContent: data}}, ...sender);
  }
  const init = sent[0].message;
  assert.equal(init.method, "ui/initialize");
  const ready = async (origin = "https://host.example.test") => {
    deliver({id: init.id, result: {hostContext: {}}}, parent, origin);
    await new Promise(resolve => setImmediate(resolve));
  };
  return {element, sent, parent, deliver, result, ready, init};
}

test("only the initialized parent and its origin may supply results or replies", async () => {
  const app = view("mail-view");
  const mail = {entry_id: "synthetic-id", subject: "Synthetic subject"};
  app.deliver({id: app.init.id, result: {}}, {});
  app.result(mail);
  assert.equal(app.element("mail").innerHTML, "");
  await app.ready();
  app.result(mail, {});
  app.result(mail, app.parent, "https://other.example.test");
  assert.equal(app.element("mail").innerHTML, "");
  const before = app.sent.length;
  app.deliver({method: "tools/call", id: 55, params: {name: "outlook_delete_mail"}});
  assert.equal(app.sent.length, before + 1);
  assert.equal(app.sent.at(-1).message.error.code, -32601);
  app.result(mail);
  assert.match(app.element("mail").innerHTML, /Synthetic subject/);
  const click = app.element("readBtn").listeners.click();
  const call = app.sent.at(-1);
  assert.equal(call.message.params.name, "outlook_mark_mail");
  assert.equal(call.target, "https://host.example.test");
  app.deliver({id: call.message.id, error: {message: "wrong sender"}}, {});
  app.deliver({id: call.message.id, result: {}});
  await click;
  assert.equal(app.element("error").innerHTML, "");
  await app.element("delBtn").listeners.click();
  const deletion = app.element("delBtn").listeners.click();
  const delCall = app.sent.at(-1).message;
  assert.equal(delCall.params.name, "outlook_delete_mail");
  app.deliver({id: delCall.id, result: {}});
  await deletion;
  const summary = app.sent.at(-1).message;
  assert.equal(summary.method, "ui/update-model-context");
  assert.ok(!JSON.stringify(summary).includes(mail.subject));
});

test("opaque parent still receives the handshake and escaped mail", async () => {
  const app = view("mail-view");
  await app.ready("null");
  const label = "<synthetic-marker> & 'quoted'";
  app.result({entry_id: "synthetic", subject: label, body: label, from: label,
              attachments: [{filename: label}]}, app.parent, "null");
  const html = app.element("mail").innerHTML;
  assert.ok(html.includes("&lt;synthetic-marker&gt; &amp; &#39;quoted&#39;"));
  assert.ok(!html.includes("<synthetic-marker>"));
  assert.ok(app.sent.every(({message, target}) => target === "*" && message.method !== "tools/call"));
});

test("calendar refresh removes previously visible details and drill-down for private slots", async () => {
  const app = view("calendar");
  await app.ready();
  const start = "2026-11-23T10:00:00", end = "2026-11-23T11:00:00";
  const subject = "<synthetic-appointment>";
  app.result({items: [{entry_id: "synthetic-id", subject, start, end}], count: 1});
  const agenda = app.element("agenda");
  assert.ok(agenda.innerHTML.includes("&lt;synthetic-appointment&gt;"));
  assert.ok(!agenda.innerHTML.includes(subject));
  assert.ok(agenda.innerHTML.includes("data-index="));
  app.result({items: [{subject: "Appointment", busy_status: 2, start, end}], count: 1});
  assert.match(agenda.innerHTML, /Appointment/);
  assert.match(agenda.innerHTML, /Busy/);
  assert.doesNotMatch(agenda.innerHTML, /synthetic|data-index=/);
  app.result({items: [], count: 0});
  assert.match(agenda.innerHTML, /No events/);
  assert.doesNotMatch(agenda.innerHTML, /Appointment|synthetic|data-index=/);
  assert.ok(app.sent.every(({message}) => message.method !== "tools/call"));
});
