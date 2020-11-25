import Vue from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'
import axios from 'axios'
import './plugins/element.js'
import VueClipboards from 'vue-clipboard2'

Vue.config.productionTip = false
axios.defaults.baseURL = 'http://localhost:9191/'
Vue.prototype.$http = axios

Vue.use(VueClipboards);

new Vue({
  router,
  store,
  render: h => h(App)
}).$mount('#app')
