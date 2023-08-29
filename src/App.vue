<template>
  <div class="file-container">
    <input type="file" @change="getFiles" />
  </div>
</template>

<script setup>
import XLSX from "xlsx";
const getFiles = (e) => {
  var files = e.target.files;
  var fileReader = new FileReader();
  fileReader.onload = (ev) => {
    try {
      var data = ev.target.result;
      var workbook = XLSX.read(data, {
        type: "binary",
      });
      console.log(workbook);
      const wsname = workbook.SheetNames[0]; // 取第一张表
      const ws = XLSX.utils.sheet_to_json(workbook.Sheets[wsname]); // 生成json表格内容
      const dataArr = [];
      ws.map((item) => {
        let obj = {
          arr: [],
        };
        console.log(item);
        Object.values(item).map((child, index) => {
          obj.arr.push("name" + index);
          obj[`name${index}`] = child;
        });
        dataArr.push(obj);
        console.log(dataArr);
      });
      console.log(ws);
    } catch (error) {
      console.log("文件类型不正确");
      return;
    }

    // var fromTo = "";
  };
  console.log(files);
  fileReader.readAsBinaryString(files[0]);
};
</script>

<style lang="scss"></style>
