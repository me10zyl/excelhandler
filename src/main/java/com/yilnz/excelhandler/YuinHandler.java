package com.yilnz.excelhandler;

import com.yilnz.surfing.core.Site;
import com.yilnz.surfing.core.SurfSpider;

import javax.sound.sampled.*;
import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class YuinHandler {

	public static void main(String[] args) throws IOException, LineUnavailableException, UnsupportedAudioFileException {
		File file = new File("/Users/zyl/Downloads/录音文件丢失第一批数据.txt");
		BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(file)));
		String str = null;
		List<String> callId = new ArrayList<>();
		while((str = br.readLine()) != null){
			Matcher matcher = Pattern.compile(",(\\w+\\.wav)").matcher(str);
			if(matcher.find()) {
				//System.out.println(matcher.group(1));
				callId.add(matcher.group(1));
			}
		}

		File out = new File(file.getParent(), file.getName() + "-录音秒数.txt");
		File err = new File(file.getParent(), file.getName() + "-录音有问题的数据.txt");
		PrintWriter pw1 = new PrintWriter(out);
		PrintWriter pw2 = new PrintWriter(err);
		for (String s : callId) {
			String url = "http://taojin-freeswitch-800.oss-cn-hangzhou.aliyuncs.com/" + s;
			System.out.println("下载：" + s);
			//File file1 = SurfSpider.downloadIfNotExist("/tmp/yuyin/" + s, url, Site.me());
			//Clip clip = AudioSystem.getClip();
			Double durationInSeconds = null;
			try {
				AudioInputStream ais = AudioSystem.getAudioInputStream(new URL(url));
				AudioFormat format = ais.getFormat();
				long frames = ais.getFrameLength();
				durationInSeconds = (frames+0.0) / format.getFrameRate();
			}catch (Exception e){
				//System.out.println(s + " " + e.getCause());
				//e.printStackTrace();
			}
			if (durationInSeconds != null) {
				String x = s.replaceAll("\\.wav", "") + " " + durationInSeconds;
				pw1.println(x);
				pw1.flush();
			}else{
				pw2.println(s);
				pw1.flush();
			}
		}
		pw1.close();
		pw2.close();

	}

}
