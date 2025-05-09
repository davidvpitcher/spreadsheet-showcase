<template>
  <div class="page">
    <h1>About This App</h1>

    <p>
      <strong>Spreadsheet Showcase</strong> is a Vue 3 application that demonstrates how to integrate
      the powerful <code>Handsontable</code> spreadsheet library into a modern, user-friendly tool.
    </p>

    <p>
      It includes a suite of formatting, data manipulation, and visualization features, allowing users to interact
      with spreadsheet data in intuitive and productive ways.
    </p>

    <h2>Key Features</h2>
    <ul>
      <li>🧪 Generate mock data</li>
      <li>🎨 Format cells with bold, italics, text/background color, alignment, and size</li>
      <li>🔲 Visual border editor with grid-style selection</li>
      <li>🧊 Freeze top row or first column (toggleable)</li>
      <li>💾 Auto-save support (toggleable) for data and formatting</li>
      <li>🔎 Search with match count and clearing</li>
      <li>📊 Column type analysis</li>
      <li>↕️ Manual column and row movement</li>
      <li>♻️ Reset all formatting and colors</li>
      <li>📤 Export CSV of output spreadsheet</li>
      <li>⚠️ Highlight error-prone columns</li>
      <li>🔍 Keyword row detection</li>
      <li>Can be built as an executable (using Electron)</li>
    </ul>

    <h2>Purpose</h2>
    <p>
      This tool serves as a reference implementation for Vue developers working with spreadsheet-style data entry and
      formatting needs. It showcases how to build a fast, extensible, and maintainable interface around Handsontable,
      with clean separation of concerns and Vue-friendly architecture.
    </p>

    <h2>Tech Stack</h2>
    <ul>
      <li>Vue 3 + Composition API</li>
      <li>Handsontable (community edition)</li>
      <li>Custom modular utilities for formatting, alignment, and persistence</li>
    </ul>

    <h2>Repository</h2>
    <p>
      You can find the source code on GitHub once the repository is published. Contributions and feedback are welcome!
    </p>
  </div>
</template>

<script>
export default {
  name: 'AboutPage'
}
</script>

<style scoped>
.page {
  max-width: 900px;
  margin: 0 auto;
  padding: 2rem 1rem;
  font-size: 16px;
  line-height: 1.6;
}

.page h1 {
  font-size: 2.4rem;
  margin-bottom: 1.2rem;
  color: #2c3e50;
}

.page h2 {
  font-size: 1.6rem;
  margin-top: 2rem;
  margin-bottom: 0.8rem;
  color: #34495e;
}

.page ul {
  margin-left: 1.5rem;
  padding-left: 1rem;
}

.page li {
  margin-bottom: 0.5rem;
}
</style>
