<template>
  <div class="flex content-center q-gutter-lg">
    <mahak-tehran/>
    <mahak-khorasan/>
  </div>
</template>

<script>
import MahakTehran from './MahakTehran'
import MahakKhorasan from './MahakKhorasan'

export default {
  name: 'Home',
  components: { MahakTehran, MahakKhorasan },
  data () {
    return {}
  },
  methods: {},
}
</script>

<style scoped>

</style>