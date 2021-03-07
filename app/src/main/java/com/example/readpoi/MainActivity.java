package com.example.readpoi;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.content.Intent;
import android.view.Menu;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;

public class MainActivity extends Activity {
	public String filePath_xlsx = Environment.getExternalStorageDirectory()
			+ "/tencent/micromsg/download/5.4-5.20订餐.xlsx";
	public String filePath_docx = Environment.getExternalStorageDirectory()
			+ "/22[1].docx";
	public String filePath_doc = Environment.getExternalStorageDirectory()
			+ "/11[1].doc";
	public String filePath_xls = Environment.getExternalStorageDirectory()
			+ "/tencent/micromsg/download/5.4-5.20订餐.xls";

	private Button btn_load_xlsx;
	private Button btn_load_doc;
	private Button btn_load_docx;
	private Button btn_load_xls;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		btn_load_xlsx = (Button) findViewById(R.id.btn_load);
		btn_load_xlsx.setOnClickListener(new OnClickListener() {

			@Override
			public void onClick(View v) {
				Intent i = new Intent();
				i.setClass(MainActivity.this, FileReadActivity.class);
				Bundle bundle = new Bundle();
				bundle.putString("filePath", filePath_xlsx);
				i.putExtras(bundle);
				startActivity(i);
			}
		});
		
		btn_load_xls = (Button) findViewById(R.id.btn_load_xls);
		btn_load_xls.setOnClickListener(new OnClickListener() {

			@Override
			public void onClick(View v) {
				Intent i = new Intent();
				i.setClass(MainActivity.this, FileReadActivity.class);
				Bundle bundle = new Bundle();
				bundle.putString("filePath", filePath_xls);
				i.putExtras(bundle);
				startActivity(i);
			}
		});
	    
		btn_load_doc = (Button) findViewById(R.id.btn_load_doc);
		btn_load_doc.setOnClickListener(new OnClickListener() {

			@Override
			public void onClick(View v) {
				Intent i = new Intent();
				i.setClass(MainActivity.this, FileReadActivity.class);
				Bundle bundle = new Bundle();
				bundle.putString("filePath", filePath_doc);
				i.putExtras(bundle);
				startActivity(i);
			}
		});
		
		btn_load_docx = (Button) findViewById(R.id.btn_load_docx);
		btn_load_docx.setOnClickListener(new OnClickListener() {

			@Override
			public void onClick(View v) {
				Intent i = new Intent();
				i.setClass(MainActivity.this, FileReadActivity.class);
				Bundle bundle = new Bundle();
				bundle.putString("filePath", filePath_docx);
				i.putExtras(bundle);
				startActivity(i);
			}
		});
		
		
	}
	
}
