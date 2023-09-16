<template>
  <div class="container">
    <h2>Product Price Finder</h2>
    <div class="form-group">
      <label for="fileInput">Upload Excel File:</label>
      <div class="file-upload">
        <input type="file"
          id="fileInput"
          @change="handleFileChange"
          accept=".xlsx" />
        <label for="fileInput"
          class="custom-file-upload">Choose File</label>
      </div>
    </div>
    <div class="form-group">
      <label for="sizeInput">Size:</label>
      <input type="text"
        id="sizeInput"
        v-model="size"
        placeholder="Size" />
    </div>
    <div class="form-group">
      <label for="dimensionsInput">Dimensions:</label>
      <input type="text"
        id="dimensionsInput"
        v-model="dimensions"
        placeholder="Dimensions" />
    </div>
    <div class="form-group">
      <label for="widthInput">Width:</label>
      <input type="text"
        id="widthInput"
        v-model="width"
        placeholder="Width" />
    </div>
    <div class="form-group">
      <button @click="findPrice">Find Price</button>
    </div>
    <p v-if="price !== null"
      class="result">Price: {{ price }}</p>
  </div>
</template>
<script setup>
import { ref } from 'vue';
import * as XLSX from 'xlsx';

const size = ref('');
const dimensions = ref('');
const width = ref('');
const price = ref(null);
let uploadedFile = null;

const handleFileChange = (event) => {
  const file = event.target.files[0];
  if (file) {
    uploadedFile = file;
  }
};

const findPrice = async () => {
  if (!uploadedFile) {
    alert('Please upload an Excel file first.');
    return;
  }

  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet['!ref']);

    // Iterate over the rows in the sheet
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      const sizeCol = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v;
      const dimensionsCol = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })]?.v;
      const widthCol = sheet[XLSX.utils.encode_cell({ r: row, c: 2 })]?.v;
      const priceCol = sheet[XLSX.utils.encode_cell({ r: row, c: 3 })]?.v;


      // Check if the values match
      if (
        sizeCol == size.value &&
        dimensionsCol == dimensions.value &&
        widthCol == width.value
      ) {
        price.value = priceCol;
        return;
      }
    }

    // If no matching row is found
    price.value = null;
    alert('No matching data found in the Excel file.');
  };

  reader.readAsArrayBuffer(uploadedFile);
};
</script>

<style>
.container {
  max-width: 400px;
  margin: 0 auto;
  padding: 20px;
  border: 1px solid #ccc;
  border-radius: 5px;
  background-color: #f9f9f9;
}

h2 {
  text-align: center;
}

.form-group {
  margin-bottom: 15px;
}

label {
  display: block;
  font-weight: bold;
}

input[type="text"] {
  width: 100%;
  padding: 10px;
  border: 1px solid #ccc;
  border-radius: 5px;
}

button {
  width: 100%;
  padding: 10px;
  background-color: #007bff;
  color: #fff;
  border: none;
  border-radius: 5px;
  cursor: pointer;
}

button:hover {
  background-color: #0056b3;
}

.result {
  margin-top: 10px;
  font-weight: bold;
}

</style>