<template>
    <q-layout view="hHh lpR fFf">

        <q-header elevated class="bg-primary text-white">
            <q-toolbar>
                <q-toolbar-title>
                    برنامه استخراج اطلاعات حساب بانک ملی بیماران
                </q-toolbar-title>
            </q-toolbar>
        </q-header>

        <q-page-container>
            <q-page padding>
                <slot></slot>
            </q-page>
        </q-page-container>

    </q-layout>
</template>

<script>
  
  export default {
    data: function () {
      return {
      }
    },
    methods: {
    }
  }
</script>