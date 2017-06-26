package com.zamahaka.exceltest

import android.os.Bundle
import android.os.Environment
import android.support.v7.app.AppCompatActivity
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream


class MainActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        findViewById(R.id.btnSheet).setOnClickListener { generateSheet(createOutputFile("test.xls")) }
        findViewById(R.id.btnDrawing).setOnClickListener { generateChart(createOutputFile("chart.xls")) }
        findViewById(R.id.btnCopy).setOnClickListener { copyFile(createOutputFile("copied.xls")) }
    }

    private fun copyFile(file: File?) {
        val inputStream = assets.open("chart.xls")
        val outputStream = FileOutputStream(file)

        inputStream.copyTo(outputStream)
        inputStream.close()
        outputStream.close()
    }

    private fun generateChart(file: File?) {
        val inputStream = assets.open("chart.xls")
        val outputStream = FileOutputStream(file)

        inputStream.copyTo(outputStream)
        inputStream.close()
        outputStream.close()

        val createdStream = FileInputStream(file)

        val workbook = HSSFWorkbook(createdStream)
        val sheet = workbook.getSheetAt(0)

        for (row in 1 until 50) {
            sheet.createRow(row).createCell(0).setCellValue(getTime(row * 6))
        }

        val editedStream = FileOutputStream(file)

        workbook.write(editedStream)

        createdStream.close()
        editedStream.close()
    }

    private fun getTime(i: Int): String {
        val rem = i.rem(60)
        val div = i.div(60)

        return "$div:$rem"
    }

    private fun generateSheet(file: File?) {

    }

    fun isExternalStorageReadOnly(): Boolean {
        val extStorageState = Environment.getExternalStorageState()
        if (Environment.MEDIA_MOUNTED_READ_ONLY == extStorageState) {
            return true
        }
        return false
    }

    fun isExternalStorageAvailable(): Boolean {
        val extStorageState = Environment.getExternalStorageState()
        if (Environment.MEDIA_MOUNTED == extStorageState) {
            return true
        }
        return false
    }

    private fun createOutputFile(name: String): File? {
        val mediaStorageDir = File(Environment.getExternalStorageDirectory().toString()
                + "/Android/data/"
                + applicationContext.packageName
                + "/Files")
        if (!mediaStorageDir.exists()) {
            if (!mediaStorageDir.mkdirs()) {
                return null
            }
        }
        val mediaFile: File
        mediaFile = File(mediaStorageDir.path + File.separator + name)
        return mediaFile
    }
}
