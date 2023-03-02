package com.jaytrairat.forensicdoc

import android.Manifest.permission.READ_EXTERNAL_STORAGE
import android.Manifest.permission.WRITE_EXTERNAL_STORAGE
import android.content.Context
import android.content.pm.PackageManager
import android.graphics.Color
import android.graphics.drawable.Drawable
import android.graphics.drawable.GradientDrawable
import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Button
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.util.*


class MainActivity : AppCompatActivity() {


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        var txtDocumentTo: TextView = findViewById(R.id.txtDocumentTo)
        var txtOriginalFrom: TextView = findViewById(R.id.txtOriginalFrom)
        var txtOriginalNumber: TextView = findViewById(R.id.txtOriginalNumber)
        var txtOriginalDate: TextView = findViewById(R.id.txtOriginalDate)
        var txtOriginalName: TextView = findViewById(R.id.txtOriginalName)
        var txtNumberOfPages: TextView = findViewById(R.id.txtNumberOfPages)

        checkPermission()
        val currentDate = SimpleDateFormat("yyyy/MM/dd", Locale.getDefault()).format(Date())

        val btnGenerateIndexPage = findViewById<Button>(R.id.btnGenerateIndexPage)
        btnGenerateIndexPage.setOnClickListener {
            createDocxFromTemplate()
        }

        val sharedPref = getSharedPreferences("prefs", Context.MODE_PRIVATE)
        txtDocumentTo.text = sharedPref.getString("txtDocumentTo", "")
        txtOriginalFrom.text = sharedPref.getString("txtOriginalFrom", "")
        txtOriginalNumber.text = sharedPref.getString("txtOriginalNumber", "")
        txtOriginalDate.text = sharedPref.getString("txtOriginalDate", currentDate)
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
            var txtOriginalDate: TextView = findViewById(R.id.txtOriginalDate)
            var txtOriginalName: TextView = findViewById(R.id.txtOriginalName)
            var txtNumberOfPages: TextView = findViewById(R.id.txtNumberOfPages)

            val documentTo = txtDocumentTo.text.toString()
            val originalFrom = txtOriginalFrom.text.toString()
            val originalNumber = txtOriginalNumber.text.toString()
            val originalDate = txtOriginalDate.text.toString()
            val originalName = txtOriginalName.text.toString()
            val numberOfPages = txtNumberOfPages.text.toString()


            val sharedPref = getSharedPreferences("prefs", Context.MODE_PRIVATE)
            with(sharedPref.edit()) {
                putString("txtDocumentTo", documentTo)
                putString("txtOriginalFrom", originalFrom)
                putString("txtOriginalNumber", originalNumber)
                putString("txtOriginalDate", originalDate)
                putString("txtOriginalName", originalName)
                putString("txtNumberOfPages", numberOfPages)
                apply()
            }

            var isError = false // to keep track of whether there is an error
            val errorBorderColor = Color.RED

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

            if (originalDate.isNullOrEmpty()) {
                txtOriginalDate.background =
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

                // Load the docx template
                val templateInputStream = resources.openRawResource(R.raw.index_case_template)
                val document = XWPFDocument(templateInputStream)

                // Replace placeholders with data from the TextView
                val data = mapOf(
                    "document_to" to documentTo,
                    "original_from" to originalFrom,
                    "original_number" to originalNumber,
                    "original_date" to originalDate,
                    "original_name" to originalName,
                    "number_of_result" to numberOfPages,
                )
                for (paragraph in document.paragraphs) {
                    for (run in paragraph.runs) {
                        var text = run.text()
                        for ((key, value) in data) {
                            text = text.replace("$key", value)
                        }
                        run.setText(text, 0)
                    }
                }

                // Save the filled template as a new docx file in the Downloads folder
                val now = LocalDateTime.now()
                val formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss")
                var index = 0
                var outputFile: File
                do {
                    val filename = "${formatter.format(now)}_${index}_page.docx"
                    outputFile = File(
                        Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS),
                        filename
                    )
                    index++
                } while (outputFile.exists())

                val fos = FileOutputStream(outputFile)
                document.write(fos)
                fos.close()

                // Close the template document
                document.close()


                Toast.makeText(this, "Document created", Toast.LENGTH_LONG).show()
            } else {
                Toast.makeText(this, "Text cannot be null", Toast.LENGTH_LONG).show()
            }


        } catch (error: Exception) {
            Log.e("ERROR", error.toString())
            Toast.makeText(this, "Failed to exports", Toast.LENGTH_LONG).show()
        }
    }
}