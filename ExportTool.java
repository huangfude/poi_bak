package ssm.system.poi;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


public class ExportTool {

    
    /**
     * 解决设置名称时的乱码  
     * @param request
     * @param fileName
     * @return
     */
    public static String processFileName(HttpServletRequest request, String fileName) {  
        String codedfilename = null;  
        try {  
            String agent = request.getHeader("USER-AGENT");  
            if (null != agent && -1 != agent.indexOf("MSIE") || null != agent  
                    && -1 != agent.indexOf("Trident")) {// ie  
  
                String name = java.net.URLEncoder.encode(fileName, "UTF8");  
  
                codedfilename = name;  
            } else if (null != agent && -1 != agent.indexOf("Mozilla")) {// 火狐,chrome等  
  
  
                codedfilename = new String(fileName.getBytes("UTF-8"), "iso-8859-1");  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return codedfilename;  
    }
    
    /**
     * 设置生成office的map
     * @param bean
     * @return
     */
    public static Map<String, String> setOfficeMap(Object bean){
        Map<String, String> params = new HashMap<String, String>();
        if( bean != null ){
            Method[] methods = bean.getClass().getDeclaredMethods();
            for(int i=0;i<methods.length;i++) {
                String methodName = methods[i].getName();
                if (methodName.substring(0,3).toUpperCase().equals("GET")) {    
                    Method method = methods[i];
                    String key = methodName.substring(3).toLowerCase();
                    try {
                        Object objValue = method.invoke(bean, new Object[]{});
                        String value = "";
                        if(objValue!=null){
                            if(objValue instanceof String){
                                value = objValue.toString();
                            } else if (objValue instanceof Integer){
                                value = objValue.toString();
                            } else if (objValue instanceof Date){
                                SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
                                value = sf.format(objValue);
                            }
                            
                        }
                        
                        params.put(key,value);  
                    } catch (Exception e) {
                        e.printStackTrace();
                    }    
                }    
            }  
        }
        return params;
    }
    
    /**
     * 导出文件
     * @param bean
     * @param specialMap
     * @param templeName
     * @param outName
     * @param request
     * @param response
     * @throws IOException
     */
    public static void exportFile(Object bean, Map<String, String> specialMap, String templeName, String outName,
            HttpServletRequest request, HttpServletResponse response) throws IOException{
        
        //转换参数
        Map<String, String> params = setOfficeMap(bean);
        //替换特殊参数
        params.putAll(specialMap);
        
        String fileNameInResource = request.getSession().getServletContext().getRealPath("/temple")+templeName;  
        MSWordTool changer = new MSWordTool();
        changer.setTemplate(fileNameInResource);
        
        changer.replaceBookMark(params);
        
        String fileName = outName;
        // 防乱码
        fileName = processFileName(request, fileName);
        response.reset(); 
        response.setContentType("application/vnd.ms-word;charset=utf-8");  
        response.setHeader("Content-Disposition", "attachment; filename=\""+ fileName + "\"");
        response.setBufferSize(1024); 
        
        OutputStream fos = response.getOutputStream();
        try {
            changer.document.write(fos);
            fos.flush();
            fos.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
    }
    
}
