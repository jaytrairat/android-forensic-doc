package com.jaytrairat.forensicdoc

import android.Manifest.permission.READ_EXTERNAL_STORAGE
import android.Manifest.permission.WRITE_EXTERNAL_STORAGE
import android.app.DownloadManager
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.graphics.Color
import android.graphics.drawable.Drawable
import android.graphics.drawable.GradientDrawable
import android.media.MediaScannerConnection
import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Button
import android.widget.DatePicker
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.*


class MainActivity : AppCompatActivity() {


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        var txtDocumentTo: TextView = findViewById(R.id.txtDocumentTo)
        var txtOriginalFrom: TextView = findViewById(R.id.txtOriginalFrom)
        var txtOriginalNumber: TextView = findViewById(R.id.txtOriginalNumber)
        var txtOriginalName: TextView = findViewById(R.id.txtOriginalName)
        var txtNumberOfPages: TextView = findViewById(R.id.txtNumberOfPages)

        checkPermission()

        val btnGenerateIndexPage = findViewById<Button>(R.id.btnGenerateIndexPage)
        btnGenerateIndexPage.setOnClickListener {
            createDocxFromTemplate()
        }


        val btnOpenDownload = findViewById<Button>(R.id.btnOpenDownload)
        btnOpenDownload.setOnClickListener {
            startActivity(Intent(DownloadManager.ACTION_VIEW_DOWNLOADS))
        }

        val sharedPref = getSharedPreferences("prefs", Context.MODE_PRIVATE)
        txtDocumentTo.text = sharedPref.getString("txtDocumentTo", "")
        txtOriginalFrom.text = sharedPref.getString("txtOriginalFrom", "")
        txtOriginalNumber.text = sharedPref.getString("txtOriginalNumber", "")
        txtOriginalName.text = sharedPref.getString("txtOriginalName", "")
        txtNumberOfPages.text = sharedPref.getString("txtNumberOfPages", "")
    }

    private fun getErrorBorderDrawable(color: Int): Drawable {
        val strokeWidth = 3 // the width of the border
        val strokeColor = color // the color of the border
        return GradientDrawable().apply {
            setStroke(strokeWidth, strokeColor)
        }
    }

    private fun checkPermission() {

        if (ContextCompat.checkSelfPermission(
                this,
                READ_EXTERNAL_STORAGE
            ) == PackageManager.PERMISSION_GRANTED &&
            ContextCompat.checkSelfPermission(
                this,
                WRITE_EXTERNAL_STORAGE
            ) == PackageManager.PERMISSION_GRANTED
        ) {
            // Permissions have been granted, do your work here
        } else {
            // Permissions have not been granted, request them
            val REQUEST_CODE = 30
            ActivityCompat.requestPermissions(
                this,
                arrayOf(READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE),
                REQUEST_CODE
            )
        }
    }

    fun createDocxFromTemplate() {
        try {
            var txtDocumentTo: TextView = findViewById(R.id.txtDocumentTo)
            var txtOriginalFrom: TextView = findViewById(R.id.txtOriginalFrom)
            var txtOriginalNumber: TextView = findViewById(R.id.txtOriginalNumber)
            var txtOriginalDate: DatePicker = findViewById(R.id.txtOriginalDate)
            var txtOriginalName: TextView = findViewById(R.id.txtOriginalName)
            var txtNumberOfPages: TextView = findViewById(R.id.txtNumberOfPages)

            val documentTo = txtDocumentTo.text.toString()
            val originalFrom = txtOriginalFrom.text.toString()
            val originalNumber = txtOriginalNumber.text.toString()
            val originalName = txtOriginalName.text.toString()
            val numberOfPages = txtNumberOfPages.text.toString()

            val originalYear = txtOriginalDate.getYear()
            val originalMonth = txtOriginalDate.getMonth()
            val originalDate = txtOriginalDate.getDayOfMonth()

            val calendar = Calendar.getInstance()
            calendar.set(originalYear, originalMonth, originalDate)

            val sharedPref = getSharedPreferences("prefs", Context.MODE_PRIVATE)
            with(sharedPref.edit()) {
                putString("txtDocumentTo", documentTo)
                putString("txtOriginalFrom", originalFrom)
                putString("txtOriginalNumber", originalNumber)
                putString("txtOriginalDate", calendar.toString())
                putString("txtOriginalName", originalName)
                putString("txtNumberOfPages", numberOfPages)
                apply()
            }

            var isError = false // to keep track of whether there is an error
            val errorBorderColor = Color.RED
            txtDocumentTo.background = null
            txtOriginalFrom.background = null
            txtOriginalNumber.background = null
            txtOriginalName.background = null
            txtNumberOfPages.background = null

            if (documentTo.isNullOrEmpty()) {
                txtDocumentTo.background =
                    getErrorBorderDrawable(errorBorderColor) // set the border to red
                isError = true // set isError to true to indicate that there is an error
            }

            if (originalFrom.isNullOrEmpty()) {
                txtOriginalFrom.background =
                    getErrorBorderDrawable(errorBorderColor) // set the border to red
                isError = true // set isError to true to indicate that there is an error
            }

            if (originalNumber.isNullOrEmpty()) {
                txtOriginalNumber.background =
                    getErrorBorderDrawable(errorBorderColor) // set the border to red
                isError = true // set isError to true to indicate that there is an error
            }

            if (originalName.isNullOrEmpty()) {
                txtOriginalName.background =
                    getErrorBorderDrawable(errorBorderColor) // set the border to red
                isError = true // set isError to true to indicate that there is an error
            }

            if (numberOfPages.isNullOrEmpty()) {
                txtNumberOfPages.background =
                    getErrorBorderDrawable(errorBorderColor) // set the border to red
                isError = true // set isError to true to indicate that there is an error
            }
            if (!isError) {
                val longThaiDateFullFormatter = SimpleDateFormat("d MMMM yyyy", Locale("th", "TH"))
                val longThaiDateFormatter = SimpleDateFormat("d MMMM yyyy", Locale("th", "TH"))
                val shortThaiDateFormatter = SimpleDateFormat("  /MMM/yy", Locale("th", "TH"))

                val currentThaiLongDate = longThaiDateFormatter.format(Date())
                val currentThaiShortDate = shortThaiDateFormatter.format(Date())
                val thaiLongDate = longThaiDateFullFormatter.format(calendar.time)

                val templateInputStream = resources.openRawResource(R.raw.index_case_template)

                val replaceParams = mapOf(
                    "documentTo" to documentTo,
                    "originalFrom" to originalFrom,
                    "originalNumber" to originalNumber,
                    "originalName" to originalName,
                    "numberOfPages" to numberOfPages,
                    "originalDate" to thaiLongDate,
                    "dateLong" to currentThaiLongDate,
                    "dateShort" to currentThaiShortDate
                )

                val dateFormat = SimpleDateFormat("yyyy_MM_dd-HH_mm_ss", Locale.US)
                val timestamp = dateFormat.format(Date())
                val exportFilename = "$timestamp-$documentTo-export.docx"

                val downloadFolder =
                    Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
                val outputFile = File(downloadFolder, exportFilename)

                XWPFDocument(templateInputStream).use { doc ->

                    for (para in doc.paragraphs) {
                        val paraText = para.text
                        if (replaceParams.keys.any { it in paraText }) {
                            for (para in doc.paragraphs) {
                                for (run in para.runs) {

                                    var text = run.text()
                                    Log.e("LINE", text)
                                    for ((key, value) in replaceParams) {
                                        text = text.replace(key, value)
                                    }
                                    run.setText(text, 0)
                                }
                            }
                        }
                    }
                    FileOutputStream(outputFile).use { outputStream ->
                        doc.write(outputStream)
                    }
                }

                // Notify the user that the file has been saved
                MediaScannerConnection.scanFile(this, arrayOf(outputFile.path), null, null)
                Toast.makeText(
                    this,
                    "File saved to ${outputFile.absolutePath}",
                    Toast.LENGTH_LONG
                ).show()

            } else {
                Toast.makeText(this, "Text cannot be null", Toast.LENGTH_LONG).show()
            }


        } catch (error: Exception) {
            Log.e("ERROR", error.toString())
            Toast.makeText(this, "Failed to exports", Toast.LENGTH_LONG).show()
        }
    }
}