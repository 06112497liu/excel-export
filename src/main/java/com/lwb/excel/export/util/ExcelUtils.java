package com.lwb.excel.export.util;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;
import com.lwb.excel.export.annotation.Export;
import com.lwb.excel.export.exception.UtilException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.Supplier;

import static com.lwb.excel.export.enums.FileType.XLSX;
import static com.lwb.excel.export.util.ExcelUtils.Constant.*;
import static com.lwb.excel.export.util.ExcelUtils.ExcelStyle.headerStyle;

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
        StackTraceElement stackTrace = getStackTrace(Export.class, Thread.currentThread().getStackTrace());
        ExcelConfig config = parseYml(stackTrace);
        // 配置完整性校验
        config.validate();
        return generateExcel(config, supplier.get());
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
                        sheet.addMergedRegion(
                            new CellRangeAddress(Integer.parseInt(index[0]),
                                Integer.parseInt(index[1]),
                                Integer.parseInt(index[2]),
                                Integer.parseInt(index[3])
                            )
                        );
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
                list.forEach(item -> {
                    config.getFields()
                        .forEach(fieldName -> {
                            SXSSFCell cell = row.createCell(row.getPhysicalNumberOfCells());
                            try {
                                cell.setCellValue(getFieldValue(item, fieldName));
                            } catch (Exception e) {
                                LOGGER.error(e.getMessage(), e);
                                throw new UtilException(e.getMessage());
                            }
                        });
                });
            });

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
     * @param obj
     * @param fieldName
     * @return
     */
    private static String getFieldValue(Object obj, String fieldName) throws NoSuchFieldException, IllegalAccessException {
        if (obj == null || StringUtils.isEmpty(fieldName)) {
            return EMPTY;
        }

        // 如果传入对象是map 直接获取key值
        if (obj instanceof Map) {
            return ((Map) obj).containsKey(fieldName) ? (String) ((Map) obj).get(fieldName) : EMPTY;
        }
        if (fieldName.contains(POINT)) {
            int i = fieldName.indexOf(POINT);
            String currentFieldName = fieldName.substring(0, i);
            String nextFieldName = fieldName.substring(i + 1, fieldName.length() - 1);
            Field field = obj.getClass().getField(currentFieldName);
            if (!field.isAccessible()) {
                field.setAccessible(true);
            }
            Object o = field.get(obj);
            if (field.isAccessible()) {
                field.setAccessible(false);
            }
            return getFieldValue(o, nextFieldName);
        } else {
            return formatFieldValue(obj, fieldName);
        }
    }

    /**
     * 格式化字段的值
     * </p>
     * 日期字段根据JsonFormat注解的样式格式化
     * @param obj
     * @param fieldName
     * @return
     */
    private static String formatFieldValue(Object obj, String fieldName) throws NoSuchFieldException, IllegalAccessException {
        Field field = obj.getClass().getField(fieldName);
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
        String fileName = String.format("%s_%s", config.getFileName(), UUID.randomUUID());

        FileOutputStream out = null;
        try {
            String fileFullPath = getFileFullPath(String.format("%s.%s", fileName, XLSX.getSuffix()));
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
     * @param trace 方法栈
     * @return
     */
    private static ExcelConfig parseYml(StackTraceElement trace) {
        try {
            Class<?> clazz = Class.forName(trace.getClassName());
            Method method = clazz.getDeclaredMethod(trace.getMethodName());
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
    private static <T extends Annotation> StackTraceElement getStackTrace(Class<T> type, StackTraceElement... stackTrace) {
        return
            Arrays.stream(stackTrace)
                .filter(trace -> {
                    try {
                        Class<?> clazz = Class.forName(trace.getClassName());
                        return
                            Optional.ofNullable(clazz.getDeclaredMethod(trace.getMethodName()))
                                .map(m -> m.getAnnotation(type) != null)
                                .orElse(false);
                    } catch (NoSuchMethodException | ClassNotFoundException e) {
                        LOGGER.error(e.getMessage(), e);
                    }
                    return false;
                })
                .findFirst()
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
    }


}
