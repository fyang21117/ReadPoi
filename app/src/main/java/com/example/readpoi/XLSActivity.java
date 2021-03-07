package com.example.readpoi;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.DocumentsContract;
import android.provider.MediaStore;
import android.support.v4.content.FileProvider;
import android.util.Log;
import android.view.View;
import android.widget.TextView;
import android.widget.Toast;

import java.io.File;
import java.util.Map;

public class XLSActivity extends Activity {

    private Map<String,Double> mMoneyMap;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_xls);
        findViewById(R.id.open).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.setType("*/*");//无类型限制
                intent.addCategory(Intent.CATEGORY_OPENABLE);
                startActivityForResult(intent, 1);
            }
        });
        findViewById(R.id.share).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if (mMoneyMap == null) {
                    Toast.makeText(XLSActivity.this, "没有汇总数据", Toast.LENGTH_SHORT).show();
                    return;
                }
                String fileName = new ExcelMaker().write(mMoneyMap);
                File file = new File(fileName);
                Uri uri =  Uri.fromFile(file);
                Intent intent = new Intent();
                intent.setAction(Intent.ACTION_SEND);
                intent.setType("*/*");
                intent.putExtra(Intent.EXTRA_STREAM, uri);
                intent = intent.createChooser(intent, "share");
                XLSActivity.this.startActivity(intent);
            }
        });
    }

    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        if (resultCode == Activity.RESULT_OK) {
            Uri uri = data.getData();
            if (Build.VERSION.SDK_INT > 19) {//4.4以后
                String path = getPath(this, uri);
                Log.e("ZXX", "文件路径：" + path);
                ((TextView)findViewById(R.id.fileName)).setText(path);
                Toast.makeText(this,path,Toast.LENGTH_SHORT).show();
                ExcelProcessor excelProcessor = new ErmaoExcelProcessor(path);
                excelProcessor.readSheet();
                mMoneyMap = ((ErmaoExcelProcessor) excelProcessor).getMoneyMap();
                String content = "";
                for (String key: mMoneyMap.keySet()){
                    Log.e("ZXX", key + "： " + mMoneyMap.get(key));
                    content += key + ":           " + mMoneyMap.get(key) + "\n";
                }
                ((TextView)findViewById(R.id.sumContent)).setText(content);
                findViewById(R.id.share).setVisibility(View.VISIBLE);
            }
        }
    }

    public String getPath(final Context context, final Uri uri) {

                 final boolean isKitKat = Build.VERSION.SDK_INT >= 19;

                 // DocumentProvider
                 if (isKitKat && DocumentsContract.isDocumentUri(context, uri)) {
                         // ExternalStorageProvider
                         if (isExternalStorageDocument(uri)) {
                                 final String docId = DocumentsContract.getDocumentId(uri);
                                 final String[] split = docId.split(":");
                                 final String type = split[0];

                                 if ("primary".equalsIgnoreCase(type)) {
                                         return Environment.getExternalStorageDirectory() + "/" + split[1];
                                     }
                             }
                         // MediaProvider
                         else if (isMediaDocument(uri)) {
                                 final String docId = DocumentsContract.getDocumentId(uri);
                                 final String[] split = docId.split(":");
                                 final String type = split[0];

                                 Uri contentUri = null;
                                 if ("image".equals(type)) {
                                         contentUri = MediaStore.Images.Media.EXTERNAL_CONTENT_URI;
                                     } else if ("video".equals(type)) {
                                         contentUri = MediaStore.Video.Media.EXTERNAL_CONTENT_URI;
                                     } else if ("audio".equals(type)) {
                                         contentUri = MediaStore.Audio.Media.EXTERNAL_CONTENT_URI;
                                     }

                                 final String selection = "_id=?";
                                 final String[] selectionArgs = new String[]{split[1]};

                                 return getDataColumn(context, contentUri, selection, selectionArgs);
                             }
                     }
                 // MediaStore (and general)
                 else if ("content".equalsIgnoreCase(uri.getScheme())) {
                         return getDataColumn(context, uri, null, null);
                     }
                 // File
                 else if ("file".equalsIgnoreCase(uri.getScheme())) {
                         return uri.getPath();
                     }
                 return null;
    }

    public String getDataColumn(Context context, Uri uri, String selection,
                                String[] selectionArgs) {

                 Cursor cursor = null;
                 final String column = "_data";
                 final String[] projection = {column};

                 try {
                         cursor = context.getContentResolver().query(uri, projection, selection, selectionArgs,
                                         null);
                         if (cursor != null && cursor.moveToFirst()) {
                                 final int column_index = cursor.getColumnIndexOrThrow(column);
                                 return cursor.getString(column_index);
                             }
                     } finally {
                         if (cursor != null)
                                 cursor.close();
                     }
                 return null;
    }

    public boolean isExternalStorageDocument(Uri uri) {
        return "com.android.externalstorage.documents".equals(uri.getAuthority());
    }

    public boolean isDownloadsDocument(Uri uri) {
        return "com.android.providers.downloads.documents".equals(uri.getAuthority());
    }

    public boolean isMediaDocument(Uri uri) {
        return "com.android.providers.media.documents".equals(uri.getAuthority());
    }

}
