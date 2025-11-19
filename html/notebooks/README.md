
# 🏗️ HTML5 Tutorial

A structured, notebook-ready introduction to HTML5, covering document structure, elements by category, semantics, media, forms, tables, and interactive components.

---

## 📜 1. HTML5 Document Structure

* Modern HTML starts with the required `<!DOCTYPE html>` declaration.
* `<html>` defines the root of the document.
* `<head>` contains metadata, script links, stylesheets, icons, and SEO information.
* `<body>` contains all visible content.
* Always specify `lang=""` for accessibility and screen readers.

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>HTML5 Example</title>
</head>
<body>
  <h1>Hello HTML5</h1>
</body>
</html>
```

---

## 🧩 2. Sectioning Elements

* Sectioning defines the logical structure of a webpage.
* Includes `<header>`, `<nav>`, `<main>`, `<section>`, `<article>`, `<aside>`, and `<footer>`.
* Helps define document outline used by search engines and screen readers.
* Use headings (`<h1>`–`<h6>`) to create hierarchy within each section.

```html
<header>
  <h1>My Website</h1>
  <nav>
    <a href="#about">About</a>
    <a href="#contact">Contact</a>
  </nav>
</header>

<main>
  <section id="about">
    <article>
      <h2>About Us</h2>
      <p>We build accessible and modern web experiences.</p>
    </article>
  </section>
</main>

<footer>
  <p>&copy; 2025 MySite</p>
</footer>
```

---

## ✏️ 3. Phrasing & Text-Level Elements

* Inline elements affect text content but do not break line flow.
* Common text semantics: `<strong>`, `<em>`, `<mark>`, `<code>`, `<abbr>`, `<time>`.
* Prefer semantic tags over purely visual ones (`<b>`, `<i>`).
* Enhance accessibility, meaning, and machine readability.

```html
<p>
  Learning <strong>HTML5</strong> is <em>fun</em>!
  Today's date is <time datetime="2025-11-09">Nov 9, 2025</time>.
  Highlight text with <mark>mark</mark>, or show code like <code>&lt;div&gt;</code>.
</p>
```

---

## 📄 4. Grouping Elements

* Group block-level content with `<div>` or structural tags.
* `<p>` wraps paragraphs of text.
* `<blockquote>` cites external sources.
* `<pre>` preserves white-space formatting (useful for code).
* `<hr>` represents thematic breaks in content.

```html
<div class="intro">
  <p>Welcome to the HTML5 tutorial.</p>

  <blockquote cite="https://developer.mozilla.org/">
    “HTML is the standard markup language for web documents.”
  </blockquote>

  <pre>
Line 1
    Line 2 (indented)
  </pre>

  <hr>
</div>
```

---

## 🖼️ 5. Embedded Content (Media)

* Handles rich content such as images, audio, video, graphics, and external frames.
* Common tags: `<img>`, `<audio>`, `<video>`, `<iframe>`, `<canvas>`, `<svg>`.
* Always include `alt` text for images to support screen readers.
* `<canvas>` supports scriptable drawing; `<svg>` supports resolution-independent graphics.

```html
<img src="logo.png" alt="Company Logo" width="120">

<video controls width="300">
  <source src="intro.mp4" type="video/mp4">
  Your browser does not support the video tag.
</video>

<audio controls>
  <source src="sound.mp3" type="audio/mpeg">
  Audio not supported.
</audio>

<canvas id="myCanvas" width="200" height="120"></canvas>
```

---

## 🧮 6. Tables

* Use tables to display structured data—not for layout.
* Compose with `<table>`, `<thead>`, `<tbody>`, `<tfoot>`, `<tr>`, `<th>`, `<td>`.
* Always add `<caption>` for accessibility and clarity.
* Use semantic row/column headers.

```html
<table>
  <caption>Monthly Budget</caption>
  <thead>
    <tr><th>Month</th><th>Income</th><th>Expenses</th></tr>
  </thead>
  <tbody>
    <tr><td>January</td><td>$4500</td><td>$3200</td></tr>
    <tr><td>February</td><td>$4600</td><td>$3100</td></tr>
  </tbody>
  <tfoot>
    <tr><td colspan="3">End of Report</td></tr>
  </tfoot>
</table>
```

---

## 📝 7. Forms & Input Controls

* Forms collect user input and submit it to a server or script.
* Common elements: `<form>`, `<label>`, `<input>`, `<textarea>`, `<select>`, `<button>`.
* Associate `<label>` with form controls using `for` and `id`.
* HTML5 adds input types like `email`, `date`, `url`, `range`, and `number`.
* Use `required`, `placeholder`, and `pattern` for validation.

```html
<form action="/submit" method="post">
  <label for="name">Name</label>
  <input id="name" name="name" type="text" required>

  <label for="email">Email</label>
  <input id="email" name="email" type="email" placeholder="you@example.com">

  <label for="topic">Topic</label>
  <select id="topic" name="topic">
    <option>HTML5</option>
    <option>CSS3</option>
  </select>

  <textarea name="message" rows="4">Enter your message…</textarea>

  <button type="submit">Send</button>
</form>
```

---

## 🌐 8. Links & Navigation

* `<a>` creates hyperlinks using `href`.
* `<nav>` organizes groups of navigation links.
* Use relative URLs for internal pages and absolute URLs for external destinations.
* Links should be clear and accessible.

```html
<nav>
  <ul>
    <li><a href="/index.html">🏠 Home</a></li>
    <li><a href="/about.html">📘 About</a></li>
    <li><a href="https://developer.mozilla.org/">🌍 MDN</a></li>
  </ul>
</nav>
```

---

## 🧭 9. Semantic & Accessibility Elements

* HTML5 introduces rich semantic elements to convey meaning.
* Examples: `<main>`, `<figure>`, `<figcaption>`, `<details>`, `<summary>`, `<dialog>`, `<address>`.
* Semantic markup improves accessibility, SEO, and maintainability.

```html
<main>
  <figure>
    <img src="architecture.jpg" alt="Modern Building">
    <figcaption>An example of modern HTML5 layout</figcaption>
  </figure>

  <address>
    Contact us at <a href="mailto:info@example.com">info@example.com</a>
  </address>
</main>
```

---

## 🧰 10. Interactive Elements (Native HTML)

* HTML5 provides interactive widgets without requiring JavaScript.
* `<details>` and `<summary>` create collapsible panels.
* `<dialog>` defines modal and non-modal dialog boxes.
* Browser support is strong and improving across modern engines.

```html
<details open>
  <summary>Click to toggle details</summary>
  <p>This feature requires no JavaScript.</p>
</details>

<dialog id="infoBox">
  <p>Hello from an HTML5 dialog!</p>
  <button onclick="infoBox.close()">Close</button>
</dialog>

<button onclick="infoBox.showModal()">Open Dialog</button>
```

---

## 🧱 11. HTML5 Graphics (Canvas & SVG)

* `<canvas>` supports low-level raster graphics via JavaScript.
* `<svg>` supports vector graphics and scales without loss of quality.
* SVG can include shapes, gradients, paths, text, filters, and animations.

```html
<!-- Simple SVG -->
<svg width="120" height="60">
  <rect width="120" height="60" fill="#3a7afe" rx="10"></rect>
  <text x="60" y="35" fill="#fff" text-anchor="middle">SVG</text>
</svg>
```

---

## 🧮 12. HTML5 Data Attributes

* Custom metadata stored directly on elements with the `data-*` prefix.
* Accessed easily from JavaScript.
* Great for annotations, state, configuration values, and dynamic behavior.

```html
<button data-role="primary" data-id="42">Click Me</button>

<script>
  const btn = document.querySelector("button");
  console.log(btn.dataset.role); // "primary"
  console.log(btn.dataset.id);   // "42"
</script>
```

---

## 🗺️ 13. HTML5 Best Practices

* Use semantic elements instead of generic `<div>`s.
* Ensure each page has one `<main>` element.
* Always include `alt` attributes on images.
* Use accessible names (`aria-label`, `aria-expanded`) when needed.
* Organize headings logically (do not skip levels).
* Prefer `<button>` over clickable `<div>`s for interactive UI.

---

# 📘 License

This tutorial may be used freely in educational, instructional, and training materials.

---

# 🙌 Contributing

Pull requests that add examples, fix accessibility issues, or expand semantic coverage are welcome.


