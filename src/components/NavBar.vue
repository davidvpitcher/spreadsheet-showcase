<template>
  <nav class="navbar">
    <router-link to="/" exact>Home</router-link>
    <router-link to="/about">About</router-link>
    <router-link to="/contact">Contact</router-link>
    <router-link to="/spreadsheetshowcase">Spreadsheet-Showcase</router-link>

  </nav>
</template>

<script>
export default {
  name: 'NavBar'
}
</script>

<style scoped>
.navbar {
  background-color: #2c3e50;
  padding: 1rem 2rem;
  display: flex;
  gap: 1.5rem;
  align-items: center;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
a {
  color: white;
  text-decoration: none;
  font-weight: 500;
}
a:hover {
  text-decoration: underline;
  color: #a2d2ff;
}
.router-link-exact-active {
  font-weight: bold;
  text-decoration: underline;
}
</style>
