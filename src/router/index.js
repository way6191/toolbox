import Vue from 'vue'
import VueRouter from 'vue-router'
import CreateEB from '@/views/CreateEB'
import Home from '@/views/Home'

Vue.use(VueRouter)

const routes = [
  {
    path: '/',
    component: Home
  },
  {
    path: '/createeb',
    component: CreateEB
  },
]

const router = new VueRouter({
  routes
})

export default router
