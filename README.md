# DOCX 转 H5 转换器（Java 8）

一个可执行 JAR 工具：接收 `docx` 文件路径参数，转换为 HTML5 页面，并尽量保留常见样式（段落、标题、加粗、斜体、下划线、字体大小、颜色、表格）。

## 构建

要求：
- JDK 8+
- Maven 3+

执行：

```bash
mvn clean package
```

生成可执行包：

- `target/docx-to-h5-converter-1.0.0-jar-with-dependencies.jar`

## 运行

```bash
java -jar target/docx-to-h5-converter-1.0.0-jar-with-dependencies.jar /path/to/input.docx
```

默认在同目录生成同名 `.html` 文件。

也可以指定输出路径：

```bash
java -jar target/docx-to-h5-converter-1.0.0-jar-with-dependencies.jar /path/to/input.docx /path/to/output.html
```

## 说明

- 已保留：段落结构、标题（Heading1~4 映射为 h1~h4）、加粗、斜体、下划线、字体颜色、字体大小、换行、表格。
- 对非常复杂的 Word 布局（浮动对象、艺术字、复杂页眉页脚、嵌套样式冲突）可能无法完全等价还原。
