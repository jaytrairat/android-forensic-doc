package com.jaytrairat.forensicdoc

import android.Manifest.permission.WRITE_EXTERNAL_STORAGE
import android.content.pm.PackageManager
import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.xmlbeans.XmlOptions
import java.io.File
import java.io.FileOutputStream
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.util.*


class MainActivity : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        val REQUEST_CODE_WRITE_EXTERNAL_STORAGE_PERMISSION = 123
        if (ContextCompat.checkSelfPermission(
                this,
                WRITE_EXTERNAL_STORAGE
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                this,
                arrayOf(WRITE_EXTERNAL_STORAGE),
                REQUEST_CODE_WRITE_EXTERNAL_STORAGE_PERMISSION
            )
        }

        val btnGenerateIndexPage = findViewById<Button>(R.id.btnGenerateIndexPage)
        btnGenerateIndexPage.setOnClickListener {
            if (ContextCompat.checkSelfPermission(
                    this,
                    WRITE_EXTERNAL_STORAGE
                ) != PackageManager.PERMISSION_GRANTED
            ) {
                ActivityCompat.requestPermissions(
                    this,
                    arrayOf(WRITE_EXTERNAL_STORAGE),
                    REQUEST_CODE_WRITE_EXTERNAL_STORAGE_PERMISSION
                )
            } else {
                // Permission has already been granted, save the file
                createDocxFromTemplate()
            }
        }
    }

    fun createDocxFromTemplate() {
        try {
            val txtDocumentTo = findViewById<TextView>(R.id.txtDocumentTo)
            val txtOriginalFrom = findViewById<TextView>(R.id.txtOriginalFrom)
            val txtOriginalNumber = findViewById<TextView>(R.id.txtOriginalNumber)
            val txtOriginalDate = findViewById<TextView>(R.id.txtOriginalDate)
            val txtOriginalName = findViewById<TextView>(R.id.txtOriginalName)
            val txtNumberOfPages = findViewById<TextView>(R.id.txtNumberOfPages)

            val documentTo = txtDocumentTo.text.toString()
            val originalFrom = txtOriginalFrom.text.toString()
            val originalNumber = txtOriginalNumber.text.toString()
            val originalDate = txtOriginalDate.text.toString()
            val originalName = txtOriginalName.text.toString()
            val numberOfPages = txtNumberOfPages.text.toString()

            // Load the docx template
            val templateInputStream = resources.openRawResource(R.raw.indexpage)
            val document = XWPFDocument(templateInputStream)

            // Replace placeholders with data from the TextView
            val data = mapOf(
                "document_to" to documentTo,
                "original_form" to originalFrom,
                "original_number" to originalNumber,
                "original_date" to originalDate,
                "original_name" to originalName,
                "number_of_result" to numberOfPages,
            )
            for (paragraph in document.paragraphs) {
                for (run in paragraph.runs) {
                    var text = run.text()
                    for ((key, value) in data) {
                        text = text.replace("{$key}", value)
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
        } catch (error: Exception) {
            Log.e("ERROR", error.toString())
            Toast.makeText(this, "Failed to exports", Toast.LENGTH_LONG).show()
        }
    }
}