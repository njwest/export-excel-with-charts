<template>
  <div id="app">
    <ExportableExcelChart />
  </div>
</template>

<script>
import ExportableExcelChart from './components/ExportableExcelChart.vue';

export default {
  name: 'App',
  components: {
    ExportableExcelChart
  }
};
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>