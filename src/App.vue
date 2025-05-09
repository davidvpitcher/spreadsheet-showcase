<template>
  <div id="app">
    <NavBar />
    <main class="page-container">
      <router-view />
    </main>
  </div>
</template>

<script>
import NavBar from './components/NavBar.vue'

export default {
  name: 'App',
  components: {
    NavBar
  }
}
</script>

<style>
body {
  margin: 0;
  font-family: 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  background-color: #f7f9fb;
  color: #333;
}

.page-container {
  max-width: 960px;
  margin: 2rem auto;
  padding: 0 1.5rem;
}

</style>
