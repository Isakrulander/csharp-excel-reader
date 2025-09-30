<template>
  <div class="excel-uploader">
    <div class="upload-card">
      <div class="card-header">
        <h2>Upload Excel File</h2>
      </div>
      
      <div class="upload-area" @click="triggerFileInput" @dragover.prevent @drop.prevent="handleDrop">
        <input ref="fileInput" type="file" @change="handleFileSelect" accept=".xlsx,.xls" style="display: none;">
        <div class="upload-icon">📊</div>
        <div class="upload-text">
          Drop Excel file here or <em>click to upload</em>
        </div>
        <div class="upload-tip">
          Only .xlsx and .xls files are supported
        </div>
      </div>

      <div v-if="loading" class="loading-section">
        <div class="progress-bar">
          <div class="progress-fill" :style="{width: progress + '%'}"></div>
        </div>
        <p>Processing your Excel file...</p>
      </div>

      <div v-if="result" class="result-section">
        <h3>File Analysis Results</h3>
        
        <div class="result-info">
          <div><strong>File Name:</strong> {{ result.fileName }}</div>
          <div><strong>Worksheet:</strong> {{ result.worksheetName }}</div>
          <div><strong>Rows:</strong> {{ result.rowCount }}</div>
          <div><strong>Columns:</strong> {{ result.columnCount }}</div>
        </div>

        <div v-if="result.data && result.data.length > 0" class="data-table">
          <h4>Data Preview (First 10 rows)</h4>
          <table class="simple-table">
            <thead>
              <tr>
                <th v-for="(value, key) in result.data[0]" :key="key">{{ key }}</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(row, index) in result.data.slice(0, 10)" :key="index">
                <td v-for="(value, key) in row" :key="key">{{ value }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import { ref } from 'vue'
import axios from 'axios'

export default {
  name: 'ExcelUploader',
  setup() {
    const loading = ref(false)
    const progress = ref(0)
    const result = ref(null)
    const fileInput = ref(null)

    const triggerFileInput = () => {
      fileInput.value.click()
    }

    const handleFileSelect = (event) => {
      const file = event.target.files[0]
      if (file) {
        uploadFile(file)
      }
    }

    const handleDrop = (event) => {
      const files = event.dataTransfer.files
      if (files.length > 0) {
        uploadFile(files[0])
      }
    }

    const uploadFile = async (file) => {
      const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
      
      if (!isExcel) {
        alert('Please upload an Excel file (.xlsx or .xls)')
        return
      }

      const isLt10M = file.size / 1024 / 1024 < 10
      if (!isLt10M) {
        alert('File size cannot exceed 10MB')
        return
      }

      loading.value = true
      progress.value = 0
      result.value = null

      // Simulate progress
      const progressTimer = setInterval(() => {
        if (progress.value < 90) {
          progress.value += Math.random() * 20
        }
      }, 300)

      try {
        const formData = new FormData()
        formData.append('file', file)

        const response = await axios.post('http://localhost:5000/api/excel/upload', formData, {
          headers: {
            'Content-Type': 'multipart/form-data'
          }
        })

        clearInterval(progressTimer)
        progress.value = 100
        loading.value = false
        result.value = response.data
        alert('Excel file processed successfully!')
      } catch (error) {
        clearInterval(progressTimer)
        loading.value = false
        
        let errorMessage = 'Upload failed'
        if (error.response && error.response.data && error.response.data.error) {
          errorMessage = error.response.data.error
        }
        
        alert(errorMessage)
      }
    }

    return {
      loading,
      progress,
      result,
      fileInput,
      triggerFileInput,
      handleFileSelect,
      handleDrop
    }
  }
}
</script>

<style scoped>
.excel-uploader {
  max-width: 800px;
  width: 100%;
}

.upload-card {
  background: rgba(255, 255, 255, 0.95);
  backdrop-filter: blur(10px);
  border-radius: 12px;
  padding: 30px;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
}

.card-header {
  text-align: center;
  margin-bottom: 30px;
}

.card-header h2 {
  font-size: 1.8rem;
  font-weight: 600;
  color: #303133;
  margin: 0;
}

.upload-area {
  border: 2px dashed #d9d9d9;
  border-radius: 8px;
  padding: 40px;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s ease;
  background: #fafafa;
}

.upload-area:hover {
  border-color: #409eff;
  background-color: #f5f7fa;
}

.upload-icon {
  font-size: 48px;
  margin-bottom: 16px;
}

.upload-text {
  color: #606266;
  font-size: 16px;
  margin-bottom: 8px;
}

.upload-tip {
  color: #909399;
  font-size: 12px;
}

.loading-section {
  margin: 30px 0;
  text-align: center;
}

.progress-bar {
  width: 100%;
  height: 8px;
  background: #f0f0f0;
  border-radius: 4px;
  overflow: hidden;
  margin-bottom: 15px;
}

.progress-fill {
  height: 100%;
  background: linear-gradient(90deg, #409eff, #36d1dc);
  transition: width 0.3s ease;
}

.loading-section p {
  margin: 0;
  color: #606266;
  font-size: 14px;
}

.result-section {
  margin-top: 30px;
}

.result-section h3 {
  color: #303133;
  margin-bottom: 20px;
  font-size: 1.4rem;
}

.result-info {
  background: #f8f9fa;
  padding: 20px;
  border-radius: 8px;
  margin-bottom: 20px;
}

.result-info div {
  margin-bottom: 8px;
  font-size: 14px;
  color: #606266;
}

.data-table h4 {
  color: #303133;
  margin-bottom: 15px;
}

.simple-table {
  width: 100%;
  border-collapse: collapse;
  border: 1px solid #ebeef5;
  border-radius: 4px;
  overflow: hidden;
}

.simple-table th,
.simple-table td {
  padding: 12px 8px;
  text-align: left;
  border-bottom: 1px solid #ebeef5;
  font-size: 13px;
}

.simple-table th {
  background: #f5f7fa;
  font-weight: 600;
  color: #909399;
}

.simple-table td {
  color: #606266;
}

.simple-table tr:hover {
  background: #f5f7fa;
}
</style>