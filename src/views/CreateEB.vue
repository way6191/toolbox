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
          <el-input v-model="form.samplePath"></el-input>
        </el-form-item>
        <el-form-item label="タイトル高さ">
          <el-slider :step="1" :max="5" show-stops v-model="form.titleHeight">
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
  import ExcelUtil from '../util/ExcelUtil';

  export default {
    data() {
      return {
        form: {
          samplePath: '',
          sample: '',
          titleHeight: 2,
          imgFolder: '',
          img: '',
          scale: 1
        }
      }
    },
    methods: {
      createExcel() {
        ExcelUtil.createExcel();
      }
    }
  }
</script>