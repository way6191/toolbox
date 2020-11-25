<template>
  <div>
    <el-card>
      <el-breadcrumb separator="/">
        <el-breadcrumb-item :to="{ path: '/' }">ホームページ</el-breadcrumb-item>
        <el-breadcrumb-item>リリース票</el-breadcrumb-item>
      </el-breadcrumb>
    </el-card>

    <el-card>
      <el-input v-model="folder">
        <el-button slot="append" icon="el-icon-folder" type="success" @click="selectFolder"></el-button>
      </el-input>
    </el-card>

    <el-card>
      <el-button type="primary" @click="getInfo">取得</el-button>
      <el-button type="primary" v-clipboard:copy="copydata">コピー</el-button>
    </el-card>

    <el-card>
      <el-table :data="releaseInfo" height="400" border style="width: 100%">
        <el-table-column prop="filePath" label="パス名" width="250">
        </el-table-column>
        <el-table-column prop="fileName" label="ファイル名">
        </el-table-column>
        <el-table-column prop="fileTime" label="更新日時" width="160">
        </el-table-column>
        <el-table-column prop="fileSize" label="サイズ" width="120">
        </el-table-column>
      </el-table>
    </el-card>
  </div>
</template>

<script>
  const remote = window.require('electron').remote
  const dialog = remote.dialog

  export default {
    data() {
      return {
        copydata: "まだ取得していない。",
        folder: "",
        releaseInfo: []
      }
    },
    methods: {
      async selectFolder() {
        let options = {};
        options.title = '資産フォルダを選択';
        options.message = '資産フォルダを選択';
        options.buttonLabel = '選択';
        options.properties = ['openDirectory'];

        const folder = await dialog.showOpenDialog(options);
        this.folder = folder.filePaths[0];
      },
      async getInfo() {
        this.copydata = "";
        this.releaseInfo.length = 0;
        let result = await this.$http.get('/getReleaseInfo', {
          params: {
            folder: this.folder
          }
        });
        let copydataArr = result.data;
        this.releaseInfo.push(...copydataArr);

        copydataArr.forEach(item => {
          this.copydata = this.copydata + item.filePath + "\t";
          this.copydata = this.copydata + item.fileName + "\t";
          this.copydata = this.copydata + item.fileTime + "\t";
          this.copydata = this.copydata + item.fileSize + "\t";

          this.copydata = this.copydata + "\n";
        });
      },
      copy() {

      }
    }
  }
</script>