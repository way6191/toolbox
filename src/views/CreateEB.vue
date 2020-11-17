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
          <el-input v-model="form.tplPath">
            <el-button slot="append" icon="el-icon-folder" type="success" @click="selectFile"></el-button>
          </el-input>
        </el-form-item>
        <el-form-item label="画像フォルダ">
          <el-input v-model="form.folder">
            <el-button slot="append" icon="el-icon-folder" type="success" @click="selectFolder"></el-button>
          </el-input>
        </el-form-item>
        <el-form-item label="画像挿入間隔">
          <el-slider :step="1" :max="5" show-stops v-model="form.shift">
          </el-slider>
        </el-form-item>
        <el-form-item label="画像倍率">
          <el-slider :step="0.1" :max="2" show-stops v-model="form.scale">
          </el-slider>
        </el-form-item>
      </el-form>
    </el-card>

    <el-card>
      <el-button type="primary" @click="createExcel">作成</el-button>
    </el-card>

  </div>
</template>

<script>
  const remote = window.require('electron').remote
  const dialog = remote.dialog

  export default {
    data() {
      return {
        form: {
          tplPath: '',
          shift: 2,
          folder: '',
          scale: 1
        }
      }
    },
    methods: {
      async createExcel() {
        let result = await this.$http.post('/createExcel', {
          tplPath: this.form.tplPath,
          shift: this.form.shift,
          folder: this.form.folder,
          scale: this.form.scale
        });

        if (result.data) {
          this.$notify({
            title: 'Tips',
            message: '作成成功。',
            type: 'success'
          });
        } else {
          this.$notify.error({
            title: 'Tips',
            message: '作成失敗。'
          });
        }
      },
      async selectFile() {
        let options = {};
        options.title = 'テンプレートを選択';
        options.message = 'テンプレートを選択';
        options.buttonLabel = '選択';
        options.properties = ['openFile'];
        options.filters = [{
          name: 'テンプレート',
          extensions: ['xlsx']
        }];

        const tplPath = await dialog.showOpenDialog(options);
        this.form.tplPath = tplPath.filePaths[0];
      },
      async selectFolder() {
        let options = {};
        options.title = '画像フォルダを選択';
        options.message = '画像フォルダを選択';
        options.buttonLabel = '選択';
        options.properties = ['openDirectory'];

        const folder = await dialog.showOpenDialog(options);
        this.form.folder = folder.filePaths[0];
      },
    }
  }
</script>