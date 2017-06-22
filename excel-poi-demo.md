```java
<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi</artifactId>
	<version>3.16</version>
</dependency>

<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi-ooxml</artifactId>
	<version>3.16</version>
</dependency>
```
### 根据EXCEL模板填充数据,注意模板使用2007Excel以上版本

```java
package com.lockbur.trackr.test.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Created by wangkun23 on 2017/6/22.
 */
public class WorkbookTest {
    public static void main(String[] args) {
        ClassPathResource testXl = new ClassPathResource("Template_zhengzhoubank_person_0001.xlsx");
        InputStream is = null;
        FileOutputStream out = null;
        try {
            //读取模板
            is = testXl.getInputStream();
            XSSFWorkbook workbook = new XSSFWorkbook(is);

            //获取工作簿
            XSSFSheet personSheet001 = workbook.getSheet("个人建档1");

            //修改值
            XSSFRow row = personSheet001.getRow(0);
            Cell cell = row.createCell(0);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("王坤");

            //保存工作簿
            out = new FileOutputStream(testXl.getFile());
            workbook.write(out);//保存Excel文件
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
                if (out != null) {
                    out.close();
                }
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
```
