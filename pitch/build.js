const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const C = {
  dark: "2E2E38",
  darker: "1A1A24",
  yellow: "FFE600",
  white: "FFFFFF",
  light: "F4F4F6",
  grey: "747480",
  text: "2E2E38",
};
const HEAD = "Arial";
const BODY = "Calibri";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Yellow circle with dark icon — the deck's repeating motif
async function iconCircle(slide, Comp, x, y, d) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: C.yellow }, line: { color: C.yellow } });
  const pad = d * 0.25;
  slide.addImage({ data: await icon(Comp, C.dark), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}

function sectionLabel(slide, num, label, dark = false) {
  slide.addText(`${num}  ${label.toUpperCase()}`, {
    x: 0.5, y: 0.35, w: 9, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true,
    color: dark ? C.yellow : C.grey, charSpacing: 2, margin: 0, isTextBox: true,
  });
}

function title(slide, text, dark = false) {
  slide.addText(text, {
    x: 0.5, y: 0.65, w: 9, h: 0.7, fontFace: HEAD, fontSize: 28, bold: true,
    color: dark ? C.white : C.text, margin: 0, isTextBox: true, valign: "top",
  });
}

function card(slide, x, y, w, h, fill = C.light) {
  slide.addShape("rect", { x, y, w, h, fill: { color: fill }, line: { color: fill } });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Why EY should still hire graduates in the age of AI";

  // ---------- 1. Title ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape("rect", { x: 6.9, y: 0, w: 3.1, h: 5.625, fill: { color: C.yellow }, line: { color: C.yellow } });
    s.addImage({ data: await icon(fa.FaUserGraduate, C.dark), x: 7.55, y: 1.75, w: 1.8, h: 1.8 });
    s.addText("PITCH", {
      x: 0.5, y: 1.0, w: 6, h: 0.3, fontFace: HEAD, fontSize: 12, bold: true, color: C.yellow,
      charSpacing: 3, margin: 0, isTextBox: true,
    });
    s.addText("Why EY should still hire graduates in the age of AI", {
      x: 0.5, y: 1.4, w: 6.0, h: 2.0, fontFace: HEAD, fontSize: 36, bold: true, color: C.white,
      margin: 0, valign: "top", isTextBox: true,
    });
    s.addText("AI should take over the tasks, not the talent pipeline.", {
      x: 0.5, y: 3.7, w: 6.0, h: 0.5, fontFace: BODY, fontSize: 18, italic: true, color: C.yellow,
      margin: 0, isTextBox: true,
    });
    s.addNotes("Introduce the pitch: our argument is that EY should keep hiring graduates even as AI takes over routine tasks.");
  }

  // ---------- 2. Our position ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    sectionLabel(s, "00", "Our position");
    title(s, "AI takes the tasks. Graduates make it valuable.");
    s.addText("EY gets the most out of AI when skilled graduates are there to:", {
      x: 0.5, y: 1.5, w: 9, h: 0.4, fontFace: BODY, fontSize: 16, color: C.text, margin: 0, isTextBox: true,
    });
    const items = [
      [fa.FaDatabase, "Feed it", "Clean, structured data that AI depends on"],
      [fa.FaSearch, "Check it", "Review, question and correct its output"],
      [fa.FaHandshake, "Turn it into client value", "Explain results and build client trust"],
    ];
    const w = 2.8, gap = 0.3, y = 2.2;
    for (let i = 0; i < items.length; i++) {
      const x = 0.5 + i * (w + gap);
      card(s, x, y, w, 2.6);
      await iconCircle(s, items[i][0], x + 0.3, y + 0.3, 0.8);
      s.addText(items[i][1], {
        x: x + 0.3, y: y + 1.2, w: w - 0.6, h: 0.6, fontFace: HEAD, fontSize: 16, bold: true,
        color: C.text, margin: 0, valign: "top", isTextBox: true,
      });
      s.addText(items[i][2], {
        x: x + 0.3, y: y + 1.85, w: w - 0.6, h: 0.6, fontFace: BODY, fontSize: 14, color: C.grey,
        margin: 0, valign: "top", isTextBox: true,
      });
    }
  }

  // ---------- 3. Innovation ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    sectionLabel(s, "01", "Innovation");
    title(s, "The graduate skillset has changed");

    // Before / after comparison
    card(s, 0.5, 1.55, 2.6, 1.55, C.light);
    s.addText([
      { text: "THEN", options: { bold: true, fontSize: 11, color: C.grey, charSpacing: 2, breakLine: true } },
      { text: "“Copy, paste and attach” juniors", options: { fontSize: 15, bold: true, color: C.text } },
    ], { x: 0.7, y: 1.7, w: 2.2, h: 1.25, fontFace: BODY, margin: 0, valign: "top", isTextBox: true });

    s.addImage({ data: await icon(fa.FaArrowRight, C.dark), x: 3.2, y: 2.15, w: 0.35, h: 0.35 });

    card(s, 3.65, 1.55, 2.6, 1.55, C.dark);
    s.addText([
      { text: "NOW", options: { bold: true, fontSize: 11, color: C.yellow, charSpacing: 2, breakLine: true } },
      { text: "Finance fundamentals + tech fluency + communication", options: { fontSize: 15, bold: true, color: C.white } },
    ], { x: 3.85, y: 1.7, w: 2.2, h: 1.25, fontFace: BODY, margin: 0, valign: "top", isTextBox: true });

    s.addText([
      { text: "Already use AI tools, data analysis and automation", options: { bullet: true, breakLine: true } },
      { text: "Trained alongside AI rather than retrained around it, so they adopt new tools faster than existing staff", options: { bullet: true } },
    ], {
      x: 0.5, y: 3.35, w: 5.75, h: 1.2, fontFace: BODY, fontSize: 14, color: C.text,
      paraSpaceAfter: 6, margin: 0, valign: "top", isTextBox: true,
    });

    // Stat callout
    card(s, 6.6, 1.55, 2.9, 3.05, C.yellow);
    s.addText("80%+", {
      x: 6.8, y: 1.7, w: 2.5, h: 0.95, fontFace: HEAD, fontSize: 54, bold: true, color: C.dark,
      margin: 0, isTextBox: true,
    });
    s.addText("of respondents to EY’s own research expect tax teams to shift from purely technical tax skills towards data, process and tech skills.", {
      x: 6.8, y: 2.65, w: 2.5, h: 1.3, fontFace: BODY, fontSize: 13, color: C.dark, margin: 0,
      valign: "top", isTextBox: true,
    });
    s.addText("Exactly what new graduates bring.", {
      x: 6.8, y: 4.05, w: 2.5, h: 0.4, fontFace: BODY, fontSize: 13, bold: true, italic: true,
      color: C.dark, margin: 0, isTextBox: true,
    });

    s.addText("Result: the graduate role isn’t disappearing. It’s being upgraded.", {
      x: 0.5, y: 4.85, w: 9, h: 0.4, fontFace: HEAD, fontSize: 15, bold: true, color: C.text,
      margin: 0, isTextBox: true,
    });
  }

  // ---------- 4. Feasibility ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    sectionLabel(s, "02", "Feasibility");
    title(s, "AI can’t fully replace graduates");
    const items = [
      [fa.FaDatabase, "AI needs clean data", "In EY’s example, inconsistently coded human data confused an AI model. Many firms still run on disconnected spreadsheets and multiple ERPs."],
      [fa.FaBalanceScale, "Accountability can’t be automated", "EY says AI decisions should be open to review, questioning and correction. That review layer needs trained people."],
      [fa.FaGavel, "Explainability may become law", "EY suggests a GDPR-style path, meaning more human oversight, not less."],
      [fa.FaHandshake, "Client relationships are human", "Partners win work through trust. Graduates are the day-to-day contact where that trust starts."],
      [fa.FaClipboardCheck, "Physical & interpersonal work", "Stock counts, site visits, chasing client teams, and spotting when answers don’t add up."],
      [fa.FaUsers, "The pipeline problem", "Every future manager, partner and MD starts as a graduate. Cut hiring and the bench is empty in 10 years."],
    ];
    const cols = 3, w = 2.8, h = 1.85, gx = 0.3, gy = 0.15, y0 = 1.45;
    for (let i = 0; i < items.length; i++) {
      const col = i % cols, row = Math.floor(i / cols);
      const x = 0.5 + col * (w + gx), y = y0 + row * (h + gy);
      const highlight = i === 5;
      card(s, x, y, w, h, highlight ? C.dark : C.light);
      await iconCircle(s, items[i][0], x + 0.2, y + 0.2, 0.5);
      s.addText(items[i][1], {
        x: x + 0.8, y: y + 0.2, w: w - 1.0, h: 0.5, fontFace: HEAD, fontSize: 12, bold: true,
        color: highlight ? C.yellow : C.text, margin: 0, valign: "middle", isTextBox: true,
      });
      s.addText(items[i][2], {
        x: x + 0.2, y: y + 0.8, w: w - 0.4, h: h - 0.9, fontFace: BODY, fontSize: 11,
        color: highlight ? C.white : C.text, margin: 0, valign: "top", isTextBox: true,
      });
    }
  }

  // ---------- 5. Collaboration: the case study ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    sectionLabel(s, "03", "Collaboration", true);
    title(s, "AI + graduates = more productivity", true);
    s.addText("EY case study: 15,000 tax transactions", {
      x: 0.5, y: 1.4, w: 9, h: 0.35, fontFace: BODY, fontSize: 15, italic: true, color: C.yellow,
      margin: 0, isTextBox: true,
    });

    // Manual vs AI comparison
    const cmp = [
      ["MANUAL", "~30 hours", "95–97% accuracy", C.darker, C.white],
      ["AI", "~5 seconds", "97% accuracy", C.yellow, C.dark],
    ];
    for (let i = 0; i < 2; i++) {
      const x = 0.5 + i * 2.35;
      card(s, x, 1.95, 2.15, 1.9, cmp[i][3]);
      s.addText(cmp[i][0], {
        x: x + 0.2, y: 2.1, w: 1.75, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, charSpacing: 2,
        color: i ? C.dark : C.yellow, margin: 0, isTextBox: true,
      });
      s.addText(cmp[i][1], {
        x: x + 0.2, y: 2.5, w: 1.9, h: 0.65, fontFace: HEAD, fontSize: 23, bold: true,
        color: cmp[i][4], margin: 0, isTextBox: true,
      });
      s.addText(cmp[i][2], {
        x: x + 0.2, y: 3.25, w: 1.75, h: 0.4, fontFace: BODY, fontSize: 14, color: cmp[i][4],
        margin: 0, isTextBox: true,
      });
    }

    // Right: key takeaways
    const pts = [
      ["50%", "less human review time. People stayed in the loop, focusing on high-value, low-confidence items."],
      ["40–70%", "of tax teams’ time goes on gathering and manipulating data. AI frees it for strategy and advisory work."],
    ];
    for (let i = 0; i < pts.length; i++) {
      const y = 1.95 + i * 1.0;
      s.addText(pts[i][0], {
        x: 5.3, y, w: 1.55, h: 0.9, fontFace: HEAD, fontSize: 22, bold: true, color: C.yellow,
        margin: 0, valign: "top", isTextBox: true,
      });
      s.addText(pts[i][1], {
        x: 6.95, y: y + 0.02, w: 2.55, h: 0.9, fontFace: BODY, fontSize: 12.5, color: C.white,
        margin: 0, valign: "top", isTextBox: true,
      });
    }
    s.addText("AI did the volume. Humans did the judgement.", {
      x: 0.5, y: 4.4, w: 9, h: 0.5, fontFace: HEAD, fontSize: 18, bold: true, color: C.white,
      margin: 0, isTextBox: true,
    });
  }

  // ---------- 6. New roles ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    sectionLabel(s, "03", "Collaboration");
    title(s, "New graduate roles are being created");
    const roles = [
      [fa.FaCheckDouble, "AI output reviewer / quality control"],
      [fa.FaTable, "Data preparation & governance"],
      [fa.FaShieldAlt, "Tech risk & AI assurance"],
      [fa.FaCogs, "Tax technology & automation"],
    ];
    const w = 2.0, gap = 0.3;
    for (let i = 0; i < roles.length; i++) {
      const x = 0.5 + i * (w + gap);
      card(s, x, 1.55, w, 1.9);
      await iconCircle(s, roles[i][0], x + (w - 0.8) / 2, 1.75, 0.8);
      s.addText(roles[i][1], {
        x: x + 0.15, y: 2.7, w: w - 0.3, h: 0.65, fontFace: HEAD, fontSize: 13, bold: true,
        color: C.text, align: "center", margin: 0, valign: "top", isTextBox: true,
      });
    }
    card(s, 0.5, 3.75, 9.0, 1.3, C.dark);
    s.addImage({ data: await icon(fa.FaChartLine, C.yellow), x: 0.8, y: 4.05, w: 0.7, h: 0.7 });
    s.addText([
      { text: "Better career development", options: { bold: true, color: C.yellow, fontSize: 16, breakLine: true } },
      { text: "Graduates move to higher-value work sooner, which builds judgement faster.", options: { color: C.white, fontSize: 14 } },
    ], { x: 1.8, y: 3.9, w: 7.5, h: 1.0, fontFace: BODY, margin: 0, valign: "middle", isTextBox: true });
  }

  // ---------- 7. Closing ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addImage({ data: await icon(fa.FaQuoteLeft, C.yellow), x: 0.5, y: 0.7, w: 0.7, h: 0.7 });
    s.addText("AI makes graduates more valuable, not less.", {
      x: 0.5, y: 1.55, w: 9, h: 1.15, fontFace: HEAD, fontSize: 32, bold: true, color: C.yellow,
      margin: 0, valign: "top", isTextBox: true,
    });
    s.addText("The firms that win will be the ones that pair AI’s speed with graduates who can check it, explain it, and turn it into client trust.", {
      x: 0.5, y: 2.9, w: 8.5, h: 1.2, fontFace: BODY, fontSize: 20, color: C.white,
      margin: 0, valign: "top", isTextBox: true,
    });
    const tags = ["Check it", "Explain it", "Build trust"];
    for (let i = 0; i < tags.length; i++) {
      const x = 0.5 + i * 1.9;
      s.addShape("roundRect", { x, y: 4.35, w: 1.7, h: 0.5, fill: { color: C.darker }, line: { color: C.yellow, width: 1 }, rectRadius: 0.25 });
      s.addText(tags[i], {
        x, y: 4.35, w: 1.7, h: 0.5, fontFace: BODY, fontSize: 13, bold: true, color: C.yellow,
        align: "center", valign: "middle", margin: 0, isTextBox: true,
      });
    }
  }

  await pres.writeFile({ fileName: __dirname + "/EY_Graduate_Hiring_Pitch.pptx" });
  console.log("done");
})();
