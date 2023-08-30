<template>
  <div class="file-container">
    <input type="file" @change="getFiles" />
    <button
      @click="generateReport"
      :disabled="nameArr.length <= 0"
      v-if="!isSync"
    >
      同步数据并下载
    </button>
    <button v-else disabled>同步中...</button>
  </div>
</template>

<script setup>
import XLSX from "xlsx";
import { ref } from "vue";
const url =
  process.env.VUE_APP_BASE_API + "/cxpt/web/enterprise/getEnterpriseList.do";
// const menuData = ref([]);
// const tableData = ref([]);
const nameArr = ref([]);
const selectArr = ref([]);
const worksheet = ref([]);
const isSync = ref(false);
// 获取表格公司名称
const getFiles = (e) => {
  var files = e.target.files;
  var fileReader = new FileReader();
  fileReader.onload = (ev) => {
    try {
      var data = ev.target.result;
      var workbook = XLSX.read(data, {
        type: "binary",
      });
      const wsname = workbook.SheetNames[0]; // 取第一张表
      const ws = XLSX.utils.sheet_to_json(workbook.Sheets[wsname]); // 生成json表格内容
      // console.log(ws);
      const dataArr = [];
      ws.map((item) => {
        let obj = {
          arr: [],
        };
        Object.values(item).map((child, index) => {
          obj.arr.push(child);
          obj[`name${index}`] = child;
        });
        dataArr.push(obj);
      });
      dataArr.forEach((item) => {
        nameArr.value.push(item.name1);
      });
      nameArr.value.splice(0, 2);
      // console.log("COMPANY=", nameArr);
      // console.log("EXCELDATA=", dataArr);
      dataArr.forEach((item) => {
        worksheet.value.push(item.arr);
      });
      // console.log("worksheet=", worksheet);
      // tableData.value = dataArr;
      // console.log("tableData=", tableData);
      // menuData.value = Object.keys(ws[3]);
      // console.log("menuData=", menuData);
    } catch (error) {
      console.log("文件类型不正确");
      return;
    }

    // var fromTo = "";
  };
  fileReader.readAsBinaryString(files[0]);
};

// 导出excel
const downloadReport = (data, name = "test", pageName = "test") => {
  if (!data.length) {
    // console.log(data);
    console.log(11);
    return;
  }
  console.log(data);
  let WorkSheet = XLSX.utils.aoa_to_sheet(data);
  // console.log(WorkSheet);
  // eslint-disable-next-line camelcase
  let new_workbook = XLSX.utils.book_new();
  // console.log(new_workbook);
  XLSX.utils.book_append_sheet(new_workbook, WorkSheet, pageName);
  XLSX.writeFile(new_workbook, `${name}.xlsx`);
};

// 同步数据并下载
const generateReport = async () => {
  isSync.value = true;
  let pageIndex = 0;
  const pageSize = 10000;
  selectArr.value = [];

  const generate = async (pageIndex) => {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
      },
      mode: "cors",
      body: `mainZZ=0&aptText=&corpCode=&areaCode=0&entName=&pageSize=${pageSize}&pageIndex=${pageIndex}`,
    });

    res.json().then((res) => {
      console.log(res);
      res.data.forEach((val) => {
        // console.log(val.corpName);
        // console.log(nameArr.value);
        if (
          nameArr.value.includes(val.corpName) &&
          val.typename == "建筑业企业（施工）"
        ) {
          selectArr.value.push(val);
        }
      });
      // console.log("selectArr=", selectArr);
      // console.log((pageIndex + 1) * pageSize);
      // console.log(res.total);
      if ((pageIndex + 1) * pageSize < res.total) {
        return generate(pageIndex + 1);
      } else {
        selectArr.value.forEach((val) => {
          worksheet.value.forEach((item) => {
            if (val.corpName == item[1]) {
              item[5] = val.mark || "/";
              item[6] = val.dengji || "无";
            }
          });
        });
        // console.log("worksheet=", worksheet);
        downloadReport(worksheet.value, "excel");
        isSync.value = false;
      }
    });
  };
  generate(pageIndex);
};
</script>

<style lang="scss">
.file-container {
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  width: 300px;
  height: 100px;
}
</style>
