# lorne_excel

Excel Utils

## 使用教程

```
 <dependency>
    <groupId>com.github.1991wangliang</groupId>
    <artifactId>lorne_excel</artifactId>
    <version>1.0.0</version>
 </dependency>

```


## 读取excel文件

```

    public static List<LSheet>  readExcel() {
		try {
			// 读取文件数据
			List<LSheet> sheets = ExcelUtils.getExcelData(new File("d://data.xls"));
			// 遍历打印
			for (LSheet s : sheets) {
				for (LRow row: s.getRows()) {
					for (String v:row.getContent()) {
						System.out.print(v + "\t\t");
					}
					System.out.println();
				}

				return sheets;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return null;

	}
```

## 写入excel文件

```

    public static void writeExcel(List<LSheet> sheets) {

		try {
			 ExcelUtils.writeExcel(new File("d://data_new.xls"),sheets);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

```