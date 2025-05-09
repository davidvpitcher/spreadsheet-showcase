import { createRouter, createWebHistory, createWebHashHistory } from 'vue-router'
import HomePage from '../views/HomePage.vue'
import AboutPage from '../views/AboutPage.vue'
import ContactPage from '../views/ContactPage.vue'
import SpreadsheetShowcase from '../views/SpreadsheetShowcase.vue'

// Detect if running inside Electron
const isElectron = window && window.process && window.process.type

const routes = [
  { path: '/', name: 'HomePage', component: HomePage },
  { path: '/about', name: 'AboutPage', component: AboutPage },
  { path: '/contact', name: 'ContactPage', component: ContactPage },
  { path: '/spreadsheetshowcase', name: 'SpreadsheetShowcase', component: SpreadsheetShowcase },
]

const router = createRouter({
  history: isElectron
    ? createWebHashHistory()
    : createWebHistory(process.env.BASE_URL),
  routes
})

export default router
