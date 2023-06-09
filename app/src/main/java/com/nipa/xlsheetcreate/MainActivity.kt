package com.nipa.xlsheetcreate

import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.provider.Settings
import android.util.Log
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.annotation.RequiresApi
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.nipa.xlsheetcreate.databinding.ActivityMainBinding
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.io.IOException
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

private const val TAG = "testCheck"
private const val STORAGE_PERMISSION_CODE = 100
class MainActivity : AppCompatActivity() {
    lateinit var binding :ActivityMainBinding
    private val storagePermissionsArray = arrayOf(
        android.Manifest.permission.READ_EXTERNAL_STORAGE,
        android.Manifest.permission.WRITE_EXTERNAL_STORAGE,
        android.Manifest.permission.CAMERA
    )
    private fun checkArrayStoragePermissions(): Boolean {
        for (permission in storagePermissionsArray) {
            if (ContextCompat.checkSelfPermission(this, permission) != PackageManager.PERMISSION_GRANTED) {
                return false
            }
        }
        return true
    }
    var permissionGrant=false
    @RequiresApi(Build.VERSION_CODES.O)
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.getRoot())

        if(checkReadWritePermission()){
            permissionGrant=true

        }else{
            requestCallPermission()
        }
        binding.btnCreate.setOnClickListener {View-> createXlSeet() }
    }
    @RequiresApi(Build.VERSION_CODES.O)
    private fun createXlSeet(){
        if(permissionGrant){
            val workbook: Workbook =createWorkbook()
            createExcelFile(workbook)
        }else{
            requestCallPermission()
        }

    }


    var xldataList=ArrayList<XlData>()
    @RequiresApi(Build.VERSION_CODES.O)
    private fun createExcelFile(ourWorkbook: Workbook) {
// https://www.section.io/engineering-education/android-excel-apachepoi/

        val folderName = "test_xl"
        val folderCreate = File("${Environment.getExternalStorageDirectory()}/$folderName")


        if (folderCreate != null && !folderCreate.exists()) {
            folderCreate.mkdirs()
        }
        val ourAppFileDirectory = folderCreate
        val formatter = DateTimeFormatter.ofPattern("MM-dd-yy-HH-mm")
        val current = LocalDateTime.now().format(formatter)
        //Create an excel file called
        val excelFile = File(ourAppFileDirectory, current+"test.xlsx")
        //  Log.d("nipaerror",excelFile.absolutePath)
        //Write a workbook to the file using a file outputstream
        var xlwrite=false
        try {
            val fileOut = FileOutputStream(excelFile)
            ourWorkbook.write(fileOut)
            fileOut.close()
            xlwrite=true
        } catch (e: FileNotFoundException) {
            xlwrite=false
            e.printStackTrace()
        } catch (e: IOException) {
            xlwrite=false
            e.printStackTrace()
        }finally {
            if(xlwrite){
                binding.tvResult.text="Xl File Create Successfully please check path "+excelFile.absolutePath
            }else{
                binding.tvResult.text="Xl File Create failure  "            }

        }
    }
    private fun createWorkbook(): Workbook {
        // Creating a workbook object from the XSSFWorkbook() class
        val myWorkBook = XSSFWorkbook()

        //Creating a sheet called "statSheet" inside the workbook and then add data to it
        val sheet: Sheet = myWorkBook.createSheet("FirstSeetExample")
        // Create header CellStyle
        // Create header CellStyle
        val headerFont: Font = myWorkBook.createFont()
        headerFont.color= IndexedColors.BLUE.index
        val cellStyle = sheet.workbook.createCellStyle()
        cellStyle.fillForegroundColor = IndexedColors.RED.getIndex()
        cellStyle.setFont(headerFont)


        addData(sheet,cellStyle)

        return myWorkBook
    }
    val columNumber=3
    private fun addData(sheet: Sheet, cellStyle: CellStyle) {
        xldataList.add(XlData("Employee Id","Name","Department","Salary"))
        xldataList.add(XlData("10001","Test Name1","Clark","10000"))
        xldataList.add(XlData("10002","Test Name2","Clark","20000"))
        xldataList.add(XlData("10003","Test Name3","Clark","30000"))
        xldataList.add(XlData("10004","Test Name4","Teacher","40000"))
        xldataList.add(XlData("10005","Test Name5","Teacher","50000"))
        xldataList.add(XlData("10006","Test Name6","Teacher","60000"))
        xldataList.add(XlData("10007","Test Name7","Teacher","70000"))
        xldataList.add(XlData("10008","Test Name8","Manager","80000"))
        var rowNum=0
        for(rowdata in xldataList){
            val row: Row =sheet.createRow(rowNum++)
            var cellNo=0
            for(i in 0..columNumber){
                val ourCell = row.createCell(i)
                if(row.rowNum==0){
                    ourCell.cellStyle=cellStyle
                }
                if(i==0){
                    ourCell?.setCellValue(rowdata.empId)
                }else if(i==1){
                    ourCell?.setCellValue(rowdata.name)
                }else if(i==2){
                    ourCell?.setCellValue(rowdata.department)
                }else if(i==3){
                    ourCell?.setCellValue(rowdata.salary)
                }

            }
        }

    }


    private fun requestCallPermission(){
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R){
            //Android is 11(R) or above
            try {

                val intent = Intent()
                intent.action = Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION
                val uri = Uri.fromParts("package", this.packageName, null)
                intent.data = uri
                sdkUpperActivityResultLauncher.launch(intent)
            }
            catch (e: Exception){
                Log.e(TAG, "error ", e)
                val intent = Intent()
                intent.action = Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION
                sdkUpperActivityResultLauncher.launch(intent)
            }
        }else{
            //for below version
            ActivityCompat.requestPermissions(this,
                arrayOf(android.Manifest.permission.WRITE_EXTERNAL_STORAGE, android.Manifest.permission.READ_EXTERNAL_STORAGE),
                STORAGE_PERMISSION_CODE
            )
        }
    }

    private val sdkUpperActivityResultLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()){

        //here we will handle the result of our intent
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R){
            //Android is 11(R) or above
            if (Environment.isExternalStorageManager()){
                //Manage External Storage Permission is granted
                Log.d(TAG, "Manage External Storage Permission is granted")

            }
            else{
                //Manage External Storage Permission is denied....
                Log.d(TAG, "Permission is denied")
                toast("Manage External Storage Permission is denied....")
            }
        }

    }

    private fun checkReadWritePermission(): Boolean{

        return if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R){
            //Android is 11(R) or above
            Environment.isExternalStorageManager()
        }
        else{
            //Permission is below 11(R)
            //  checkBelowPermissionGranted()
            checkArrayStoragePermissions()
        }
    }


    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<out String>,
        grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == STORAGE_PERMISSION_CODE){
            if (grantResults.isNotEmpty() && grantResults.all { it == PackageManager.PERMISSION_GRANTED }){
                Log.d(TAG, "External Storage Permission granted")
                permissionGrant=true

            }
            else{
                //External Storage Permission denied...
                Log.d(TAG, "Some  Permission denied...")
                toast("Some Storage Permission denied...")
            }
        }
    }


    private fun toast(message: String){
        Toast.makeText(this, message, Toast.LENGTH_SHORT).show()
    }
    data class XlData(val empId:String,val name:String,val department:String,val salary:String)
}