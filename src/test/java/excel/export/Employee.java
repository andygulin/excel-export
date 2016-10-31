package excel.export;

import java.util.Date;

public class Employee {

	@Export(name = "序号")
	private Integer id;
	@Export(name = "姓名")
	private String name;
	@Export
	private Boolean sex;
	@Export
	private Date birthday;

	public Employee(Integer id, String name, Boolean sex, Date birthday) {
		this.id = id;
		this.name = name;
		this.sex = sex;
		this.birthday = birthday;
	}

	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Boolean getSex() {
		return sex;
	}

	public void setSex(Boolean sex) {
		this.sex = sex;
	}

	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}
}