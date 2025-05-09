<template>
  <div class="page">
    <h1>Welcome to the Spreadsheet Showcase</h1>

    <p>
      This is a feature-rich spreadsheet tool built with <strong>Vue 3</strong> and
      <strong>Handsontable</strong>. It demonstrates how to implement an interactive, customizable
      spreadsheet experience in a modern web app.
    </p>

    <p>
      You can explore formatting, data generation, search, column analysis, and more — all through a clean UI
      with keyboard shortcuts and persistent customization.
    </p>

    <h2>Quick Links</h2>
    <ul>
      <li><router-link to="/spreadsheetshowcase">📄 Open the Spreadsheet Showcase</router-link></li>
      <li><router-link to="/about">ℹ️ About This App</router-link></li>
    </ul>
  </div>
</template>

<script>
export default {
  name: 'HomePage'
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

.page a {
  color: #3498db;
  text-decoration: none;
}

.page a:hover {
  text-decoration: underline;
}
</style>
