package com.bzc.example.demo.utils;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import lombok.extern.slf4j.Slf4j;
import sun.misc.BASE64Encoder;

import java.beans.PropertyDescriptor;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

@Slf4j
public class FreeMarkerUtils<T> {
    private static String ftlpath="/CommonDemo/src/main/resources/templates";
    private static final String DOCUMENT_HOME = "../document/"; // 生成后的文档存放路径

    private static final String IMAGE_REGEX = "\\w+\\.(jpg|gif|bmp|png|jpeg)"; // 图片格式的文件

    private T target;

    public FreeMarkerUtils(T target) {
        this.target = target;
    }

    /**
     * 读取模版，在指定的目录中生成文件
     * @param dataMap 需要填入的数据
     * @param templateName 模版名称
     * @param fileName 文件名称
     * @throws IOException 抛出异常
     */
    public static void generateReports(Map dataMap, String templateName, String fileName) throws IOException {
        Configuration configuration = new Configuration(Configuration.VERSION_2_3_22); // 创建配置实例 
        configuration.setDefaultEncoding("UTF-8");
//        Resource resource = new ClassPathResource("template/"); // ftl 模板文件统一放至 template 包下面
//        String directory = resource.getURL().getPath();
        File lj = new File("..");
        String directory =lj.getCanonicalPath()+ftlpath;
        configuration.setDirectoryForTemplateLoading(new File(directory)); // 模板指定的路径
        Template template = configuration.getTemplate(templateName); // 获取模板 
        String filepath = DOCUMENT_HOME + File.separator + fileName;
        File outFile = new File(filepath); // 输出文件
        // 如果输出目标文件夹不存在，则创建
        if (!outFile.getParentFile().exists()) {
            if (outFile.getParentFile().mkdirs()) {
                log.info("创建目录成功！！！");
            }
        }
        // 将模板和数据模型合并生成文件 
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), StandardCharsets.UTF_8));
        try {
            template.process(dataMap, out); // 生成文件
            log.info("done !!!");
            out.flush(); // 关闭流
            out.close();
        } catch (TemplateException e) {
            log.error(e.getMessage(), e);
        }
    }

    /**
     * 获取图片编码后的值
     * @param imageUrl 本地图片路径
     * @return 图片的 base64 编码
     */
    private static String getImageString(String imageUrl) throws IOException {
        InputStream inputStream;
        byte[] data;
        inputStream = new FileInputStream(imageUrl); // 读取本地图片
        data = new byte[inputStream.available()];
        int read = inputStream.read(data);
        log.info("read " + read);
        inputStream.close();
        BASE64Encoder encoder = new BASE64Encoder();
        return encoder.encode(data);
    }

    /**
     * 图片数据转化成模板需要的值
     * @return 可以写入模板的键值对
     */
    public static Map<String, Object> parseImages(String listName,String imageName ) throws IOException {
        Map<String, Object> dataMap = new HashMap<>();
        List<String> images = new ArrayList<>();
        File lj = new File("..");
        String directory =lj.getCanonicalPath()+ftlpath;
        File[] dir = new File(directory).listFiles();
        if (null != dir && dir.length > 0) {
            if (Pattern.matches(IMAGE_REGEX,imageName)) {
                images.add(getImageString(directory + imageName));
            }
        }
        dataMap.put(listName, images);
        return dataMap;
    }

    @SuppressWarnings("unchecked")
    private Map<String, Object> handleProperties(Field field, Object obj) throws Exception {
        Map<String, Object> dataMap = new HashMap<>(); // 用于组装word页面需要的数据
        ParameterizedType parameterizedType = (ParameterizedType) field.getGenericType(); // 泛型的类型
        Class parameterTypeClazz = (Class) parameterizedType.getActualTypeArguments()[0];
        Field[] parameterizedTypeFields = parameterTypeClazz.getDeclaredFields(); // 泛型类型的所有属性
        if (null != field.get(obj)) {
            Class listClazz = field.get(obj).getClass(); // 获取 List 属性的 Class 对象
            Method getSizeOfList = listClazz.getDeclaredMethod("size"); // list 的 size 方法
            int size = (Integer) getSizeOfList.invoke(field.get(obj)); // 集合的长度
            Method getIndexOfList = listClazz.getDeclaredMethod("get", int.class); // list 集合的 get(i) 方法
            for (int index = 0; index < size; index ++) {
                Object parameterInstance = getIndexOfList.invoke(field.get(obj), index); // 泛型对象的实例
                // 遍历泛型的实例
                for (Field subsetField : parameterizedTypeFields) {
                    if (List.class.isAssignableFrom(subsetField.getType())) { // 如果该对象属性还有 list 类型
                        dataMap.put("next", handleProperties(subsetField, parameterInstance));
                    } else {
                        PropertyDescriptor propertyDescriptor = new PropertyDescriptor(subsetField.getName(), parameterInstance.getClass());
                        dataMap.put(subsetField.getName(), propertyDescriptor.getReadMethod().invoke(parameterInstance));
                    }
                }
            }
        }
        return dataMap;
    }

    /**
     * 将简单对象处理成模板需要的键值对
     * @return 可以写入模板的键值对
     */
    public Map<String, Object> parseSingleBean(T obj) throws Exception {
        Map<String, Object> dataMap = new HashMap<>(); // 用于组装word页面需要的数据
        Field[] fields = target.getClass().getDeclaredFields(); // javabean 的属性
        for (Field field : fields) {
            field.setAccessible(true);
            if (List.class.isAssignableFrom(field.getType())) { // 如果是集合类型
                dataMap.put("backup", this.handleProperties(field, obj));
            } else {
                PropertyDescriptor propertyDescriptor = new PropertyDescriptor(field.getName(), target.getClass());
                dataMap.put(field.getName(), propertyDescriptor.getReadMethod().invoke(obj));
            }
        }
        return dataMap;
    }

    /**
     * 将集合转化为模板需要的值
     * @param collectionName 集合名称
     * @param data 数据
     * @return 可以写入模板的键值对
     */
    public Map<String, Object> parseCollection(String collectionName, List<T> data) throws Exception {
        List<Map<String, Object>> collectionData = new ArrayList<>();
        for (T target : data) {
            collectionData.add(this.parseSingleBean(target));
        }
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put(collectionName, collectionData);
        return dataMap;
    }

    /**
     * 循环生成多个表格
     * @param contextName 集合命名
     * @param data 数据
     * @return map 包含每个子表格的小 map
     * @throws Exception 抛出异常
     */
    public Map<String, Object> parseComplexCollection(String contextName, List<List<T>> data, String titleName, List<T> titles) throws Exception {
        List<Map<String, Object>> collectionData = new ArrayList<>();
        // 如果存在标题
        if (null != titles && !StringUtils.isNullOrEmptyString(titleName)) {
            List<Map<String, Object>> results = new ArrayList<>();
            for (int i = 0; i < titles.size(); i ++) {
                Map<String, Object> retVal = this.parseSingleBean(titles.get(i));
                List<Map<String, Object>> contentList = new ArrayList<>();
                for (T target : data.get(i)) {
                    contentList.add(this.parseSingleBean(target));
                }
                retVal.put("content", contentList);
                results.add(retVal);
            }
            Map<String, Object> middleMap = new HashMap<>();
            middleMap.put(titleName, results);
            collectionData.add(middleMap);
        } else {
            for (List<T> outer : data) {
                collectionData.add(this.parseCollection(contextName, outer)); // 单独的每一张表格
            }
        }
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put(contextName, collectionData);
        return dataMap;
    }

}