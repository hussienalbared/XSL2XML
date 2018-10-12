
public class ExcelProperties {
private String columnName;
private String tagName;
private boolean hashed;
public String getColumnName() {
	return columnName;
}
public void setColumnName(String columnName) {
	this.columnName = columnName;
}
public String getTagName() {
	return tagName;
}
public void setTagName(String tagName) {
	this.tagName = tagName;
}
public boolean isHashed() {
	return hashed;
}
public void setHashed(boolean hashed) {
	this.hashed = hashed;
}
@Override
public String toString() {
	return "ExcelProperties [columnName=" + columnName + ", tagName=" + tagName + ", hashed=" + hashed + "]";
}

}
