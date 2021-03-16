<template>
    <div class="flex content-center items-center justify-center q-gutter-lg" style="width: 100%">
        <q-card bordered class="" style="">
            <q-card-section>
                <div class="flex items-center q-gutter-md">
                    <div>
                        <img height="70px" src="/assets/mahak_khorasan.jpg" alt="لوگوی محک"/>
                    </div>
                    <div class="text-h6">استخراج اطلاعات با فایل اکسل و فایل اکسس</div>
                </div>
            </q-card-section>

            <q-separator inset/>

            <q-card-section class="text-center q-gutter-md">
                <input type="file" ref="file" @change="handleFileUpload" style="display: none"/>
                <div>
                    <q-btn :loading="loading" color="primary" rounded label="بارگذاری فایل اکسل" icon="cloud_upload"
                           @click="$refs.file.click()"/>
                </div>

                <div>
                    <q-linear-progress v-if="loading" class="q-my-lg" stripe rounded size="25px"
                                       :value="progressPercent"
                                       color="accent">
                        <div class="absolute-full flex flex-center">
                            <q-badge color="white" text-color="accent" :label="`${progressValue} از ${progressTotal}`"/>
                        </div>
                    </q-linear-progress>
                </div>

                <div>
                    <q-btn v-if="!!downloadUrl" color="positive" rounded label="دانلود نتیجه" icon="cloud_download"
                           @click="downloadResult"/>
                </div>
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
        this.downloadUrl = null
        this.loading = true
        let timer = setInterval(async () => {
          let progress = await axios.get('/api/bank/progress')
          if (progress.data) {
            this.progressPercent = progress.data.percent / 100.0
            this.progressValue = progress.data.value
            this.progressTotal = progress.data.total
          }
        }, 700)

        let res = await axios.post('/api/bank/upload2',
          formData,
          {
            headers: {
              'Content-Type': 'multipart/form-data',
            },
          },
        )
        this.$refs.file.value = ''
        if (res.data.error) {
          this.$q.notify({
            type: 'negative',
            message: res.data.error
          })
        } else if (res.data.url) {
          this.$q.notify({
            type: 'positive',
            message: 'استخراج اطلاعات با موفقیت انجام شد.'
          })
          this.downloadUrl = res.data.url
        }
        clearInterval(timer)
        this.loading = false
        this.$emit('upload-finished')
      },

      downloadResult () {
        window.open(this.downloadUrl, '_blank')
      }
    },
  }
</script>

<style scoped>

</style>