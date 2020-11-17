<template>
  <div>
    <el-card>
      <el-breadcrumb separator="/">
        <el-breadcrumb-item :to="{ path: '/' }">ホームページ</el-breadcrumb-item>
        <el-breadcrumb-item>エビデンス作成</el-breadcrumb-item>
      </el-breadcrumb>
    </el-card>

    <el-card>
      <el-form :model="form" label-position="right" label-width="110px">
        <el-form-item label="テンプレート">
          <el-input v-model="form.tplPath"></el-input>
        </el-form-item>
        <el-form-item label="上に間隔行数">
          <el-slider :step="1" :max="5" show-stops v-model="form.shift">
          </el-slider>
        </el-form-item>
        <el-form-item label="画像フォルダー">
          <el-input v-model="form.imgFolder" webkitdirectory></el-input>
        </el-form-item>
        <el-form-item label="画像倍率">
          <el-slider :step="0.1" :max="2" show-stops v-model="form.scale">
          </el-slider>
        </el-form-item>
      </el-form>
    </el-card>

    <el-card>
      <el-button type="primary" @click="createExcel">生成</el-button>
    </el-card>

  </div>
</template>

<script>
  export default {
    data() {
      return {
        form: {
          tplPath: '',
          shift: 2,
          imgFolder: '',
          scale: 1
        }
      }
    },
    methods: {
      createExcel(){
        this.$http.post('/createExcel', {
          tplPath: this.form.tplPath,
          shift: this.form.shift,
          imgFolder: this.form.imgFolder,
          scale: this.form.scale
        });
      }
    }
  }
</script>