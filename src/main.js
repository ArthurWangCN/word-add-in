import Vue from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'

Vue.config.productionTip = false

// 必须先进行 Office 初始化
Office.onReady(() => {
  new Vue({
    router,
    store,
    render: h => h(App)
  }).$mount('#app')
})
