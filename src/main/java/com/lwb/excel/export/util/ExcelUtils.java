package com.lwb.excel.export.util;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;
import com.lwb.excel.export.annotation.Export;
import com.lwb.excel.export.exception.UtilException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.*;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URL;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.Supplier;
import java.util.stream.Stream;

import static com.lwb.excel.export.enums.FileType.XLSX;
import static com.lwb.excel.export.util.ExcelUtils.Constant.*;
import static com.lwb.excel.export.util.ExcelUtils.ExcelStyle.headerStyle;
import static com.lwb.excel.export.util.ExcelUtils.ExcelStyle.setCellRangeAddress;
import static com.lwb.excel.export.util.ExcelUtils.Headers.USER_AGENT;
import static com.lwb.excel.export.util.ExcelUtils.MediaType.APPLICATION_OCTET_STREAM_VALUE;

/**
 * @author liuweibo
 * @date 2019/8/14
 */
public class ExcelUtils {
    /**
     * 临时excel文件存放位置
     */
    private static final String TEMP_EXCEL_PATH = "temp";
    public static final String CLASSPATH_URL_PREFIX = "classpath:";
    private static Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);

    private static ThreadPoolExecutor EXECUTOR;

    static {
        int coreSize = Runtime.getRuntime().availableProcessors();
        EXECUTOR = new ThreadPoolExecutor(
            coreSize,
            coreSize << 1,
            200,
            TimeUnit.SECONDS,
            new LinkedBlockingQueue<>(coreSize << 2),
            new ThreadPoolExecutor.CallerRunsPolicy()
        );
    }

    /**
     * 格式化下载文件名函数
     */
    private static Function<String, String> FORMAT_FILE_NAME = (fileName) ->
        String.format("attachment; filename=\"%s\"", fileName);

    /**
     * 判断是否是ie内核浏览器断言
     */
    private static Predicate<String> IS_IE = userAgent ->
        // 是否是ie浏览器
        userAgent.contains("MSIE")
            // 是否是edge浏览器
            || userAgent.contains("EDGE")
            // 是否是ie内核浏览器
            || userAgent.contains("TRIDENT");

    /**
     * 默认的临时文件夹
     * @return
     */
    public static String getTempExcelPath() {
        return TEMP_EXCEL_PATH;
    }

    /**
     * 生成excel，用于后续导出
     * @param supplier 获取数据的方法
     * @return 文件名
     */
    public static String excel(Supplier<List<?>> supplier) {
        Method method = getMethod(Export.class, Thread.currentThread().getStackTrace());
        ExcelConfig config = parseYml(method);
        // 配置完整性校验
        config.validate();
        return generateExcel(config, supplier.get());
    }

    /**
     * 下载生成的临时文件
     * @param supplier 需要写入excel的数据
     */
    public static void download(Supplier<List<?>> supplier, HttpServletResponse response, HttpServletRequest request) throws UnsupportedEncodingException {
        String fileName = excel(supplier);
        ServletOutputStream out = null;
        FileInputStream in = null;
        String fileFullName = null;

        // 设置下载文件名
        String newFileName =
            Optional.of(request.getHeader(USER_AGENT).toUpperCase())
                .filter(IS_IE)
                .map(t -> {
                    try {
                        return URLEncoder.encode(fileName, UTF_8);
                    } catch (UnsupportedEncodingException e) {
                        return EMPTY;
                    }
                }).orElse(new String(fileName.getBytes(UTF_8), ISO_8859_1));

        response.setContentType(APPLICATION_OCTET_STREAM_VALUE);
        response.setHeader(Headers.CONTENT_DISPOSITION, FORMAT_FILE_NAME.apply(newFileName));

        byte[] buffer = new byte[1024];
        try {
            fileFullName = getFileFullPath(fileName);
            in = new FileInputStream(fileFullName);
            out = response.getOutputStream();
            int len;
            while ((len = in.read(buffer)) > 0) {
                out.write(buffer, 0, len);
                out.flush();
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new UtilException(e.getMessage());
        } finally {
            IOUtils.closeQuietly(in);
            IOUtils.closeQuietly(out);

            // 异步删除文件
            String name = fileFullName;
            EXECUTOR.execute(() ->
                Optional.of(new File(name))
                    // 文件是否存在
                    .filter(File::exists)
                    // 删除文件
                    .filter(File::delete)
                    .ifPresent(file -> LOGGER.debug(String.format("file %s deleted!", name)))
            );
        }
    }

    /**
     * 生成excel
     * @param config
     * @param data
     * @return 文件名
     */
    private static String generateExcel(ExcelConfig config, List<?> data) {
        SXSSFWorkbook book = new SXSSFWorkbook();
        SXSSFSheet sheet = book.createSheet();
        // 表头样式
        CellStyle headerStyle = headerStyle(book);
        List<CellRangeAddress> cellRangeAddresses = new ArrayList<>();
        // 绘制表头
        config.getHeaders()
            .forEach(headers -> {
                // 获取有记录的行数（最后有数据的行是第n行，前面有m行是空行没数据，则返回n-m）
                SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                headers.forEach(header -> {
                    String name = header.getName();
                    String merge = header.getMergeIndex();
                    Cell cell = row.createCell(row.getPhysicalNumberOfCells());

                    // 是否合并单元格
                    if (merge != null) {
                        String[] index = merge.split(",");
                        CellRangeAddress cellAddresses = new CellRangeAddress(
                            Integer.parseInt(index[0]),
                            Integer.parseInt(index[1]),
                            Integer.parseInt(index[2]),
                            Integer.parseInt(index[3])
                        );
                        sheet.addMergedRegion(cellAddresses);
                        // 收集合并的单元格，用于后续设置合并后的样式，防止合并后单元格样式丢失
                        cellRangeAddresses.add(cellAddresses);
                    }
                    cell.setCellValue(name);
                    cell.setCellStyle(headerStyle);
                });
            });

        // excel设置单元格值
        Optional.ofNullable(data)
            .filter(CollectionUtils::isNotEmpty)
            .ifPresent(list -> {
                SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                list.forEach(item -> config.getFields()
                    .forEach(fieldName -> {
                        SXSSFCell cell = row.createCell(row.getPhysicalNumberOfCells());
                        try {
                            cell.setCellValue(getFieldValue(item, fieldName));
                        } catch (Exception e) {
                            LOGGER.error(e.getMessage(), e);
                            throw new UtilException(e.getMessage());
                        }
                    }));
            });

        // 设置合并单元格后的单元格样式
        setCellRangeAddress(cellRangeAddresses, sheet);

        // 冻结表头
        Optional.ofNullable(config.getFreezePaneIndex())
            .filter(StringUtils::isNotEmpty)
            .filter(s -> s.contains(COMMA))
            .map(s -> {
                String[] index = s.split(COMMA);
                sheet.createFreezePane(
                    Integer.parseInt(index[0]),
                    Integer.parseInt(index[1]),
                    Integer.parseInt(index[2]),
                    Integer.parseInt(index[3])
                );
                return EMPTY;
            })
            // 默认冻结表头行数
            .orElseGet(() -> {
                sheet.createFreezePane(0, config.getHeaders().size());
                return EMPTY;
            });
        return save(book, config);
    }

    /**
     * 获取字段值
     * @param obj       对象
     * @param fieldName 字段名称
     * @return 字段值，转换成了String
     */
    private static String getFieldValue(Object obj, String fieldName) throws NoSuchFieldException, IllegalAccessException {
        if (obj == null || StringUtils.isEmpty(fieldName)) {
            return EMPTY;
        }
        // 如果传入对象是map 直接获取key值
        if (obj instanceof Map) {
            return ((Map) obj).containsKey(fieldName) ? (String) ((Map) obj).get(fieldName) : EMPTY;
        }
        // 支持获取嵌套对象的值（例如：user.role.name，表示获取user对象中嵌套对象role的name字段的值）
        if (fieldName.contains(POINT)) {
            int i = fieldName.indexOf(POINT);
            String currentFieldName = fieldName.substring(0, i);
            String nextFieldName = fieldName.substring(i + 1, fieldName.length());
            Field field = obj.getClass().getDeclaredField(currentFieldName);
            if (!field.isAccessible()) {
                field.setAccessible(true);
            }
            Object o = field.get(obj);
            if (field.isAccessible()) {
                field.setAccessible(false);
            }
            // 当前字段为null，不在向下获取值
            if (o == null) {
                return EMPTY;
            }
            return getFieldValue(o, nextFieldName);
        } else {
            return formatFieldValue(obj, fieldName);
        }
    }

    /**
     * 格式化字段的值
     * </p>
     * 日期字段根据JsonFormat注解的样式格式化，没有设置则使用相关默认的格式
     * @param obj       对象
     * @param fieldName 字段名称
     * @return 格式化后的值
     */
    private static String formatFieldValue(Object obj, String fieldName) throws NoSuchFieldException, IllegalAccessException {
        Field field = obj.getClass().getDeclaredField(fieldName);
        if (!field.isAccessible()) {
            field.setAccessible(true);
        }
        Object o = field.get(obj);
        if (field.isAccessible()) {
            field.setAccessible(false);
        }

        return Optional.of(o)
            .filter(ExcelUtils::isDate)
            .map(d -> {
                String pattern = Optional.ofNullable(field.getAnnotation(JsonFormat.class))
                    .map(JsonFormat::pattern)
                    .orElseGet(() -> {
                        try {
                            PropertyDescriptor descriptor = new PropertyDescriptor(fieldName, o.getClass());
                            return Optional.ofNullable(descriptor.getWriteMethod())
                                .map(m -> m.getAnnotation(JsonFormat.class))
                                .map(JsonFormat::pattern)
                                .orElse(null);
                        } catch (IntrospectionException e) {
                            LOGGER.error(e.getMessage(), e);
                            return null;
                        }
                    });
                return dateFormat(o, pattern);
            })
            .orElse(String.valueOf(o));
    }

    /**
     * 对象是不是日期对象
     * @param obj
     * @return
     */
    private static boolean isDate(Object obj) {
        return (obj instanceof Date) ||
            (obj instanceof LocalDateTime) ||
            (obj instanceof LocalDate) ||
            (obj instanceof LocalTime);
    }

    /**
     * 转换日期格式
     * @param date
     * @param pattern
     * @return
     */
    private static String dateFormat(Object date, String pattern) {
        if (date instanceof Date) {
            return DateFormatUtils.format((Date) date, pattern);
        } else if (date instanceof LocalTime) {
            return
                DateTimeFormatter.ofPattern(pattern == null ? HH_MM_SS : pattern).format((LocalTime) date);
        } else if (date instanceof LocalDate) {
            return
                DateTimeFormatter.ofPattern(pattern == null ? YYYY_MM_DD : pattern).format((LocalDate) date);
        } else if (date instanceof LocalDateTime) {
            return
                DateTimeFormatter.ofPattern(pattern == null ? YYYY_MM_DD_HH_MM_SS : pattern).format((LocalDateTime) date);
        }
        return String.valueOf(date);
    }

    /**
     * 生成临时文件，供后续下载
     * @param book
     * @param config
     */
    private static String save(Workbook book, ExcelConfig config) {
        // 生成唯一文件名
        String fileName = String.format("%s_%s.%s", config.getFileName(), UUID.randomUUID(), XLSX.getSuffix());

        FileOutputStream out = null;
        try {
            String fileFullPath = getFileFullPath(fileName);
            File file = new File(fileFullPath);
            // 创建临时文件夹
            if (!file.getParentFile().exists()) {
                file.getParentFile().mkdir();
            }
            // 创建临时文件
            if (!file.exists()) {
                file.createNewFile();
            }
            out = new FileOutputStream(file);
            book.write(out);
        } catch (IOException e) {
            LOGGER.error(e.getMessage(), e);
            throw new UtilException(e.getMessage());
        } finally {
            IOUtils.closeQuietly(out);
        }
        return fileName;
    }

    /**
     * 获得临时文件全路径
     * @param fileName 临时文件名包含后缀
     * @return 全路径文件名
     */
    private static String getFileFullPath(String fileName) throws FileNotFoundException {
        return getClassPathURL() + File.separator + TEMP_EXCEL_PATH + File.separator + fileName;
    }

    /**
     * 获取classpath路径
     * @return
     * @throws FileNotFoundException
     */
    private static String getClassPathURL() throws FileNotFoundException {
        String path = CLASSPATH_URL_PREFIX.substring(CLASSPATH_URL_PREFIX.length());
        ClassLoader cl = getDefaultClassLoader();
        URL url = (cl != null ? cl.getResource(path) : ClassLoader.getSystemResource(path));
        if (url == null) {
            String description = "class path resource [" + path + "]";
            throw new FileNotFoundException(description +
                " cannot be resolved to URL because it does not exist");
        }
        return url.getPath();
    }

    /**
     * 获取默认的类加载器
     * @return 类加载器
     */
    private static ClassLoader getDefaultClassLoader() {
        ClassLoader cl = null;
        try {
            cl = Thread.currentThread().getContextClassLoader();
        } catch (Throwable ex) {
        }
        if (cl == null) {
            cl = ExcelUtils.class.getClassLoader();
            if (cl == null) {
                try {
                    cl = ClassLoader.getSystemClassLoader();
                } catch (Throwable ex) {
                }
            }
        }
        return cl;
    }


    /**
     * 解析yml文件
     * </p>
     * 解析成ExcelConfig，用于后续初始化excel
     * @param method 被某个注解标记的方法
     * @return
     */
    private static ExcelConfig parseYml(Method method) {
        try {
            Class<?> clazz = method.getDeclaringClass();
            Export exportConfig = method.getAnnotation(Export.class);
            ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
            return mapper.readValue(
                clazz.getResourceAsStream(exportConfig.value()),
                ExcelConfig.class
            );
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new UtilException(e.getMessage());
        }
    }

    /**
     * 获取获取特定方法栈信息
     * </p>
     * 方法栈中，找到被某个注解标记的方法
     * @param type       注解类型
     * @param stackTrace 方法调用栈
     * @return 方法栈
     */
    private static <T extends Annotation> Method getMethod(Class<T> type, StackTraceElement... stackTrace) {
        return
            Arrays.stream(stackTrace)
                .map(trace -> {
                    try {
                        Class<?> clazz = Class.forName(trace.getClassName());
                        return
                            Optional.ofNullable(clazz.getDeclaredMethods())
                                .filter(ArrayUtils::isNotEmpty)
                                .map(methods -> Stream.of(methods)
                                    .filter(method -> method.getAnnotation(type) != null)
                                    .findAny()
                                    .orElse(null)
                                )
                                .orElse(null);
                    } catch (ClassNotFoundException e) {
                        LOGGER.error(e.getMessage(), e);
                    }
                    return null;
                })
                .filter(Objects::nonNull)
                .findAny()
                .orElseThrow(() -> new UtilException("没有找到ExportConfig标记的方法!"));
    }

    /**
     * excel样式
     */
    static class ExcelStyle {

        /**
         * excel头部样式
         * @param book
         * @return
         */
        public static CellStyle headerStyle(SXSSFWorkbook book) {
            CellStyle style = book.createCellStyle();

            // 基本样式
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 单元格水平方向样式
            style.setAlignment(HorizontalAlignment.CENTER);
            // 单元格垂直方向样式
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            // 自动换行
            style.setWrapText(true);

            // 设置字体
            Font font = book.createFont();
            font.setBold(true);
            style.setFont(font);

            return style;
        }

        /**
         * 设置合并单元格后的样式
         * @param addresses 合并的单元格
         * @param sheet     所属sheet
         */
        public static void setCellRangeAddress(List<CellRangeAddress> addresses, Sheet sheet) {
            addresses.forEach(address -> {
                RegionUtil.setBorderBottom(BorderStyle.THIN, address, sheet);
                RegionUtil.setBorderLeft(BorderStyle.THIN, address, sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, address, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, address, sheet);
            });
        }
    }

    /**
     * 常用常量
     */
    interface Constant {
        String EMPTY = "";
        String COMMA = ",";
        String POINT = ".";
        String YYYY_MM_DD_HH_MM_SS = "yyyy-MM-dd HH:mm:ss";
        String YYYY_MM_DD = "yyyy-MM-dd";
        String HH_MM_SS = "HH:mm:ss";
        String UTF_8 = "UTF-8";
        String ISO_8859_1 = "ISO8859_1";
    }

    interface MediaType {
        String APPLICATION_OCTET_STREAM_VALUE = "application/octet-stream";
    }

    interface Headers {
        String CONTENT_DISPOSITION = "Content-Disposition";
        String USER_AGENT = "User-Agent";
    }


}
