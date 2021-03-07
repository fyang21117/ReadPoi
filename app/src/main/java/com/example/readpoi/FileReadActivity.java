package com.example.readpoi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.TableIterator;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;
import android.util.DisplayMetrics;
import android.view.Gravity;
import android.view.WindowManager;
import android.webkit.WebView;
import android.widget.FrameLayout;
import android.widget.LinearLayout;
import android.widget.RelativeLayout;
import android.widget.TextView;

public class FileReadActivity extends Activity {
	public String nameStr = null;
	public Range range = null;
	public HWPFDocument hwpf = null;
	public String htmlPath;
	public String picturePath;
	public WebView view;
	public List pictures;
	public TableIterator tableIterator;
	public int presentPicture = 0;
	public int screenWidth;
	public FileOutputStream output;
	public File myFile;
	StringBuffer lsb = new StringBuffer();
    FR fr=null;
    TextView tv;
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.load_view_activity);
//		DisplayMetrics display = new DisplayMetrics();
//		WindowManager windowManager = this.getWindowManager();
//		windowManager.getDefaultDisplay().getMetrics(display); // ��ȡ��Ļ����
//		RelativeLayout rlFileRead=(RelativeLayout)findViewById(R.id.rl_fileread);
//		FrameLayout.LayoutParams params = new FrameLayout.LayoutParams(
//				(int) (display.widthPixels * 0.8),
//				(int) (display.heightPixels * 0.9));
//		params.gravity = Gravity.CENTER;
//		rlFileRead.setLayoutParams(params);		
		this.view = (WebView)findViewById(R.id.wv_view);
		this.view.getSettings().setBuiltInZoomControls(true); 
		this.view.getSettings().setUseWideViewPort(true);
		this.view.getSettings().setSupportZoom(true);
		this.screenWidth = this.getWindowManager().getDefaultDisplay().getWidth() - 10;//���ÿ��Ϊ��Ļ���-10
		
	//	read();
		Intent intent = this.getIntent();
		Bundle bundle = intent.getExtras();
				nameStr = bundle.getString("filePath");
			
		fr=new FR(nameStr);
	//	tv.setText(fr.returnPath);
		this.view.loadUrl(fr.returnPath);
	}



}
