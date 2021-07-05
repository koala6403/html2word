package com.ysh.exam.controller;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.util.ClassUtils;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.util.HtmlUtils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.ysh.exam.util.CustomXWPFDocument;
import com.ysh.exam.util.OfficeUtil;

@RequestMapping("/paper")
@RestController
public class GeneratePaperController {
	
	// 临时文件目录
	private String tempPath = "";
	// 模板文件目录
	private String templatePath = "";
	
	// 模板文件名
	private String templateFileName = "";
	
	@RequestMapping("download")
	public Object downloadPaper(@RequestBody String requestData, 
			HttpServletResponse response, HttpServletRequest request) {
		
		JSONObject obj = new JSONObject(requestData);
		
		// 获取临时文件夹
		String path = ClassUtils.getDefaultClassLoader().getResource("").getPath();
		this.tempPath = (path + "static/temp/").substring(1);
		this.templatePath = (path + "static/template/").substring(1);
		// 纸张大小+打印方向+装订线
		this.templateFileName = obj.getString("size") + "_" + obj.getString("direct") + "_" + obj.getString("col") +
				("1".equals(obj.get("line")) ? "_L" : "") + ".dotx";
		
		// 创建文件夹
		File tempFold = new File(tempPath);
		if (!tempFold.exists()) {
			tempFold.mkdirs();
		}
		
		// 写入Word文档
		String outputFileName = writeWordFile(obj.getString("paperContent"));
		
        InputStream is= null;
        try {
            // 清空输出流
            response.reset();
            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document ;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + outputFileName);
            response.setHeader("fileName", URLEncoder.encode(obj.getString("fileName"),"UTF-8")  + ".docx");
            
            // 读取流
            is = new FileInputStream(new File(outputFileName));

            // 写入响应里
            byte[] data = new byte[1024];
            int len = 0;
            while((len = is.read(data)) > 0) {
            	response.getOutputStream().write(data, 0, len);
            }
            
            response.getOutputStream().flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (Exception e) {
            	e.printStackTrace();
            }
        }
		
		return null;
	}
	
	/**
	 * 写入Word文档
	 * @param content HTML字符串
	 */
	private String writeWordFile(String content) {
		String finalFileName = "";
		
		Map<String, Object> param = new HashMap<String, Object>();

		// 需要删除的文件列表
		List<String> needDeleteFiles = new ArrayList<String>();
		
		// 临时HTML文件写入流
		FileOutputStream tempHtmlFos = null;
		// Word文档对象
		CustomXWPFDocument doc = null;
		// Word文档写入流
		FileOutputStream outputDocxFos = null;
		
		try {
			// 转换成标准HTML格式
			content = HtmlUtils.htmlUnescape(content);
			
			Document contentDoc = Jsoup.parse(content);

			// 遍历u标签
			Elements uList = contentDoc.select("u");
			for (Element u : uList) {
				if (u.selectFirst("span") != null) {
					// span标签添加字体
					String style = u.selectFirst("span").attr("style");
					if (!style.contains("font-family")) {
						style = "font-family: Calibri;" + style;
						u.selectFirst("span").attr("style", style);
					}
				} else {
					// u标签内没有span的情况
					String ustr = u.html();
					u.html("");
					u.append("<span style='font-family: Calibri;font-size:10.5pt'>" + ustr + "</span>");
				}
				
			}
			content = contentDoc.outerHtml();
			System.out.println(contentDoc.outerHtml());
			
			// 抽出需要替换的图片和公式
			HashMap<String, List<HashMap<String, String>>> repalceMap = getReplaceStr(content);
			// 图片公式
			List<HashMap<String, String>> maths = repalceMap.get("imgmaths");
			int count = 0;
			for (HashMap<String, String> math : maths) {
				count++;
				String key = "${mathReplace" + count + "}";
				content = content.replace(math.get("tag"), key + "\r\n");
				
				Map<String, Object> header = new HashMap<String, Object>();
				header.put("content", math.get("content"));
				header.put("mapType", "math");

				param.put(key, header);
			}

			// mathml公式
			String regex = "(<math)(.*?)(</math>)";
			Pattern r = Pattern.compile(regex, Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
			Matcher m = r.matcher(content);
	        while (m.find()) {
	        	count++;
				String key = "${mathReplace" + count + "}";
				String replaceStr = m.group(0);
				content = content.replace(replaceStr, key);
				
				Map<String, Object> header = new HashMap<String, Object>();
				header.put("content", replaceStr);
				header.put("mapType", "math");
				
				param.put(key, header);
	        }
			
			// 图片
			List<HashMap<String, String>> imgs = repalceMap.get("pics");
			count = 0;
			for (HashMap<String, String> img : imgs) {
				count++;
				//处理替换以“/>”结尾的img标签
				content = content.replace(img.get("img"), "${imgReplace" + count + "}");
				//处理替换以“>”结尾的img标签
				content = content.replace(img.get("img1"), "${imgReplace" + count + "}");
				Map<String, Object> header = new HashMap<String, Object>();

				// 找到对应文件
				String[] sep = img.get("src").replaceAll("/", "\\\\").split("\\\\");
				String imagePath = tempPath + sep[sep.length - 1];
				
				//如果没有宽高属性，默认设置为400*300
				if(img.get("width") == null || "".equals(img.get("width")) || img.get("height") == null || "".equals(img.get("height"))) {
					header.put("width", 400);
					header.put("height", 300);
				}else {
					header.put("width", (int) (Double.parseDouble(img.get("width"))));
					header.put("height", (int) (Double.parseDouble(img.get("height"))));
				}
				header.put("type", img.get("type"));
				header.put("mapType", "img");
				header.put("content", OfficeUtil.inputStream2ByteArray(new FileInputStream(imagePath), true));
				
				param.put("${imgReplace" + count + "}", header);
				needDeleteFiles.add(imagePath);
			}
			
			// 写入HTML临时文件
			tempHtmlFos = new FileOutputStream(tempPath + "temp.html");
			tempHtmlFos.write(content.getBytes("utf-8"));
			tempHtmlFos.close();
			needDeleteFiles.add(tempPath + "temp.html");
			
			// HTML转Word(JACOB)
			htmlToWord(tempPath + "temp.html", tempPath + "temp.docx");
			needDeleteFiles.add(tempPath + "temp.docx");
			
			// 临时文件（手动改好的docx文件）
			doc = OfficeUtil.generateWord(param, tempPath + "temp.docx");
			
			//最终生成的带图片的word文件
			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSS");
			finalFileName = tempPath + sdf.format(new Date()) + ".docx";
			outputDocxFos = new FileOutputStream(finalFileName);
			doc.write(outputDocxFos);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				// Word文档
				if (doc != null) {
					doc.close();
				}
				
				// Word文档写入流
				if (outputDocxFos != null) {
					outputDocxFos.flush();
					outputDocxFos.close();
				}
				
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
		
		// 删除临时文件
//		for(String delPath : needDeleteFiles) {
//			File delFile = new File(delPath);
//			delFile.delete();
//		}
		
		return finalFileName;
	}
	
	/**
	 * HTML转WORD
	 * @param html HTML文件地址
	 * @param wordFile Word文件地址
	 */
	private void htmlToWord(String html, String wordFile) {
		ActiveXComponent app = new ActiveXComponent("Word.Application"); // 启动word
		try {
			app.setProperty("Visible", new Variant(false));
			Dispatch wordDoc = app.getProperty("Documents").toDispatch();
			wordDoc = Dispatch.invoke(wordDoc, "Open", Dispatch.Method, new Object[]{templatePath + templateFileName, 
					new Variant(true), new Variant(true)}, new int[1]).toDispatch();
			Dispatch selection = app.getProperty("Selection").toDispatch();
			Dispatch.call(selection, "EndKey", new Variant(6));
			Dispatch.invoke(selection, "InsertFile", Dispatch.Method,
					new Object[] { html, "", new Variant(false), new Variant(false), new Variant(false) }, new int[3]);
			Dispatch.invoke(wordDoc, "SaveAs", Dispatch.Method, new Object[] { wordFile, new Variant(16) },
					new int[1]);
			Dispatch.call(wordDoc, "Close", new Variant(false));
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			app.invoke("Quit", new Variant[] {});
		}
	}
	
	/**
	 * 获取需要替换的内容（图片、公式）
	 * @param htmlStr 标准化后的HTML字符串
	 * @return 图片公式信息
	 * @throws Exception
	 */
	private HashMap<String, List<HashMap<String, String>>> getReplaceStr(String htmlStr) throws Exception {
		HashMap<String, List<HashMap<String, String>>> replaceParams = new HashMap<String, List<HashMap<String, String>>>();
		
		// 图片
		List<HashMap<String, String>> pics = new ArrayList<HashMap<String, String>>();
		// 图片公式
		List<HashMap<String, String>> imgmaths = new ArrayList<HashMap<String, String>>();

		Document doc = Jsoup.parse(htmlStr);
		
		// 遍历img标签
		Elements imgs = doc.select("img");
		for (Element img : imgs) {
			HashMap<String, String> map = new HashMap<String, String>();
			if (img.hasAttr("data-mathml")) {
				// mathtype格式的公式
				String mathml = img.attr("data-mathml");
				mathml = mathml.replace("«", "<").replace("»", ">").replace("¨", "\"")
						.replace("§#177;", "&plusmn;").replace("xmlns=\"http://www.w3.org/1998/Math/MathML\"", "");
				map.put("tag", img.toString());
				map.put("content", mathml);
				imgmaths.add(map);
			} else {
				// 图片
				String src = img.attr("src");
				String type = "png";
				
				// 保存图片
				String fileName = "";
				if (src.startsWith("data:image")) {
					type = src.substring(src.indexOf("/") + 1, src.indexOf(";"));
					String regex = "data:image/(png|gif|jpg|jpeg|bmp|tif|psd|ICO);base64,";
					fileName = saveBase64Img(src.replaceAll(regex, ""), type);
				} else {
					type = src.substring(src.lastIndexOf(".") + 1);
					fileName = saveNetImg(src);
				}

				if(!"".equals(img.attr("width"))) {
					map.put("width", img.attr("width"));
				}
				if(!"".equals(img.attr("height"))) {
					map.put("height", img.attr("height"));
				}
				map.put("img", img.toString().substring(0, img.toString().length() - 1) + "/>");
				map.put("img1", img.toString());
				map.put("src", img.attr("src"));
				map.put("fileName", fileName);
				map.put("type", type);
				
				pics.add(map);
			}
		}
		
		replaceParams.put("pics", pics);
		replaceParams.put("imgmaths", imgmaths);

		return replaceParams;
	}
	
	/**
	 * 保存BASE64格式的图片
	 * @param src img标签中的src属性值
	 * @param type 文件类型
	 * @return 保存的文件名
	 * @throws Exception
	 */
	private String saveBase64Img(String src, String type) throws Exception {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSS.");
		String fileName = sdf.format(new Date()) + type;
        byte[] data = src.getBytes("utf-8");  
        //new一个文件对象用来保存图片，默认保存当前工程根目录 
		String imagePath = tempPath + fileName;
        File imageFile = new File(imagePath);  
        //创建输出流  
        FileOutputStream outStream = new FileOutputStream(imageFile);  
        //写入数据  
        outStream.write(data);  
        //关闭输出流  
        outStream.close(); 
        
        return fileName;
	}
	
	/**
	 * 保存图片
	 * @param src img标签中的src属性值
	 * @return 保存的文件名
	 * @throws Exception
	 */
	private String saveNetImg(String src) throws Exception {
		//new一个URL对象  
        URL url = new URL(src);  
        //打开链接  
        HttpURLConnection conn = (HttpURLConnection)url.openConnection();  
        //设置请求方式为"GET"  
        conn.setRequestMethod("GET");  
        //超时响应时间为5秒  
        conn.setConnectTimeout(5 * 1000);  
        //通过输入流获取图片数据  
        InputStream inStream = conn.getInputStream();  
        //得到图片的二进制数据，以二进制封装得到数据，具有通用性  
        byte[] data = readInputStream(inStream);  
        //new一个文件对象用来保存图片，默认保存当前工程根目录  
		String[] sep = src.replaceAll("/", "\\\\").split("\\\\");
		String fileName = sep[sep.length - 1];
		String imagePath = tempPath + fileName;
        File imageFile = new File(imagePath);  
        //创建输出流  
        FileOutputStream outStream = new FileOutputStream(imageFile);  
        //写入数据  
        outStream.write(data);  
        //关闭输出流  
        outStream.close();
        
        return fileName;
	}
	
	/**
	 * 图片文件流转成byte数组
	 * @param inStream 图片文件流
	 * @return 转换后的byte数组
	 * @throws Exception
	 */
	private byte[] readInputStream(InputStream inStream) throws Exception{  
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();  
        //创建一个Buffer字符串  
        byte[] buffer = new byte[1024];  
        //每次读取的字符串长度，如果为-1，代表全部读取完毕  
        int len = 0;  
        //使用一个输入流从buffer里把数据读取出来  
        while( (len=inStream.read(buffer)) != -1 ){  
            //用输出流往buffer里写入数据，中间参数代表从哪个位置开始读，len代表读取的长度  
            outStream.write(buffer, 0, len);  
        }  
        //关闭输入流  
        inStream.close();  
        //把outStream里的数据写入内存  
        return outStream.toByteArray();  
    } 
}
