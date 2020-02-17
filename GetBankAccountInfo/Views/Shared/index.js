import Vue from 'vue'
import Quasar from 'quasar'
import Layout from './Layout'

Quasar.lang.set(Quasar.lang.faIr)

new Vue({
    el: '#app',
    components: {
        Layout,
        ...((typeof $components === 'undefined') ? {} : $components)
    }
})