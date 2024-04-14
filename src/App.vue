<script setup lang="ts">
import { read, utils, writeFileXLSX } from "xlsx";
import { ref } from "vue";
import type { UploadFile } from "element-plus";

type IData = Record<string, string | number>;
const data = ref<IData[]>([]);

const columns = ref<string[]>([]);

const selectA = ref<string[]>([]);

const selectB = ref<string[]>([]);

const productField = "__EMPTY";

const fields = ref<string[]>([]);

enum AverageEnum {
  averageA = "averageA",
  averageB = "averageB",
  averageDiff = "averageDiff",
}

const averageFields = [
  AverageEnum.averageA,
  AverageEnum.averageB,
  AverageEnum.averageDiff,
];

const getAverage = () => {
  columns.value = [...columns.value, ...averageFields];
  data.value.map((item) => {
    item[AverageEnum.averageA] = Object.keys(item)
      .filter((keyItem) => {
        return selectA.value.includes(keyItem);
      })
      .reduce((sum, current: string) => {
        return sum + Math.round(+item[current] * 10000);
      }, 0);
    item[AverageEnum.averageA] =
      (
        Math.round(item[AverageEnum.averageA] / selectA.value.length) / 100
      ).toFixed(2) + "%";
    item[AverageEnum.averageB] = Object.keys(item)
      .filter((keyItem) => {
        return selectB.value.includes(keyItem);
      })
      .reduce((sum, current: string) => {
        return sum + Math.round(+item[current] * 10000);
      }, 0);
    item[AverageEnum.averageB] =
      (
        Math.round(item[AverageEnum.averageB] / selectB.value.length) / 100
      ).toFixed(2) + "%";
    item[AverageEnum.averageDiff] =
      (
        ((Number.isNaN(item[AverageEnum.averageA])
          ? 0
          : parseFloat(item[AverageEnum.averageA]) * 100) -
          (Number.isNaN(item[AverageEnum.averageB])
            ? 0
            : parseFloat(item[AverageEnum.averageB]) * 100)) /
        100
      ).toFixed(2) + "%";
  });
  data.value = data.value.sort(
    (a, b) =>
      parseFloat(b[AverageEnum.averageDiff] as string) -
      parseFloat(a[AverageEnum.averageDiff] as string)
  );
};

const beforeUpload = (file: UploadFile) => {
  const reader = new FileReader();
  if (file.raw) {
    reader.readAsArrayBuffer(file.raw);
    reader.onloadend = (event) => {
      if (event.type === "loadend") {
        const workbox = read(reader.result, { dense: false });
        const { Sheets, SheetNames } = workbox;

        const sheetname = SheetNames[0];
        if (sheetname && Sheets[sheetname]) {
          const jsonData: IData[] = utils.sheet_to_json(Sheets[sheetname]);

          columns.value = jsonData[0]
            ? Object.keys(jsonData[0]).map((item) => {
                return item === productField ? "product" : item;
              })
            : [];
          fields.value = columns.value.filter((item) => item !== "product");
          data.value = jsonData;
        }
      }
    };
  }
};

const download = () => {
  const wb = utils.book_new();
  const ws = utils.json_to_sheet(data.value);
  utils.book_append_sheet(wb, ws, "Data");
  writeFileXLSX(wb, "excel-compute-average.xlsx");
};

const getDisplayNum = (val: number, field: string) => {
  return averageFields.includes(field as AverageEnum)
    ? val
    : Math.round(val * 10000) / 100 + "%";
};
</script>

<template>
  <div>
    <el-row>
      <el-upload
        :auto-upload="false"
        accept=".xls,.xlsx"
        class="upload-demo"
        :on-change="beforeUpload"
      >
        <el-button type="primary">Click to upload</el-button>
      </el-upload>
    </el-row>
    <el-row>
      <el-table border :data="data" style="height: 500px">
        <el-table-column
          :label="item"
          v-for="(item, index) in columns"
          :key="index"
          :property="item"
          :fixed="index >= columns.length - 3 ? 'right' : ''"
        >
          <template #default="scope">
            {{
              item !== "product"
                ? getDisplayNum(scope.row[item], item)
                : scope.row[productField]
            }}
          </template>
        </el-table-column>
      </el-table>
    </el-row>
    <el-row>
      <label class="label"> 月份： </label>
      <el-select
        v-model="selectA"
        placeholder="Select"
        style="width: 240px"
        multiple
      >
        <el-option v-for="(item, index) in fields" :key="index" :value="item">{{
          item
        }}</el-option>
      </el-select>
      <label class="label"> 月份： </label>
      <el-select
        v-model="selectB"
        placeholder="Select"
        style="width: 240px"
        multiple
      >
        <el-option v-for="(item, index) in fields" :key="index" :value="item">{{
          item
        }}</el-option>
      </el-select>
      <el-button type="primary" style="margin-left: 20px" @click="getAverage"
        >compute</el-button
      >
      <el-button type="primary" style="margin-left: 20px" @click="download"
        >download</el-button
      >
    </el-row>
  </div>
</template>

<style scoped></style>
