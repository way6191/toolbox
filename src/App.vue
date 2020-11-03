<template>
  <div id="app">
    <el-container>
      <el-aside width="200px">
        <div class="boxname">
          TOOL BOX
        </div>
        <el-menu router>
          <el-menu-item index="createeb">
            <span slot="title">エビ作成</span>
          </el-menu-item>
        </el-menu>
      </el-aside>
      <el-main>
        <router-view></router-view>
      </el-main>
    </el-container>
  </div>
</template>

<script>
  export default {

  }
</script>

<style>
  html,
  body,
  #app,
  .el-container {
    height: 100%;
    margin: 0;
    padding: 0;
  }

  .el-aside {
    color: #409EFF;
    text-align: center;
  }

  .boxname {
    padding: 20px;
    border: 5px solid #409EFF;
    border-right: 0;
  }

  .el-main {
    background-color: #FFFFFF;
    color: #333;
    border: 10px solid #409EFF;
    border-bottom: 0;
  }
</style>
