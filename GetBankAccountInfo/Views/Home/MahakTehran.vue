<template>
    <div class="flex content-center q-gutter-lg">
        <q-card bordered class="" style="">
            <q-card-section>
                <div class="flex items-center q-gutter-md">
                    <div>
                        <img height="70px" src="/assets/mahak_tehran.png" alt="لوگوی محک"/>
                    </div>
                    <div class="text-h6">استخراج با ساختار تهران</div>
                </div>
            </q-card-section>

            <q-separator inset/>

            <q-card-section class="text-center">
                <input type="file" ref="file" @change="handleFileUpload" style="display: none"/>
                <q-btn :loading="loading" color="primary" rounded label="بارگذاری فایل اکسل" icon="cloud_upload"
                       @click="$refs.file.click()"/>

                <q-linear-progress v-if="loading" class="q-my-lg" stripe rounded size="25px" :value="progressPercent"
                                   color="accent">
                    <div class="absolute-full flex flex-center">
                        <q-badge color="white" text-color="accent" :label="`${progressValue} از ${progressTotal}`"/>
                    </div>
                </q-linear-progress>

                <q-btn v-if="!!downloadUrl" color="secondary" rounded label="دانلود نتیجه"/>
            </q-card-section>
        </q-card>
    </div>
</template>

<script>
  export default {
    name: 'MahakTehran',
    components: {},
    data () {
      return {
        loading: false,
        progressPercent: 0.0,
        progressValue: 0,
        progressTotal: 0,
        downloadUrl: null,
      }
    },
    methods: {
      async handleFileUpload () {
        let formData = new FormData()
        formData.append('file', this.$refs.file.files[0])
        this.$emit('upload-started')
        this.progressPercent = 0.0
        this.progressValue = 0
        this.progressTotal = 0
        this.loading = true
        let timer = setInterval(async () => {
          let progress = await axios.get('/api/bank/progress')
          this.progressPercent = progress.percent
          this.progressValue = progress.value
          this.progressTotal = progress.total
        }, 1000)
        
        await axios.post('/api/bank/upload-1',
          formData,
          {
            headers: {
              'Content-Type': 'multipart/form-data',
            },
          },
        )
        clearInterval(timer)
        this.loading = false
        this.$emit('upload-finished')
      },
    },
  }
</script>

<style scoped>

</style>