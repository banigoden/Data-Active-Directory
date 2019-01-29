public class Employee {

    private String name;
    private String department;
    private String title;
    private String email;
    private String number;
    private String mobile;


   public Employee(String  name,String department, String title, String email, String number, String mobile){


        this.name = name;
        this.title = title;
        this.department = department;
        this.email = email;
        this.number = number;
        this.mobile = mobile;
    }

        public String getMobile () {

             return mobile;
        }

        public String getNumber () {

             return number;
        }

        public String getEmail() {

            return email;
        }

        public String getTitle () {

            return title;
        }

        public String getDepartment () {

            return department;
        }

        public String getName () {

             return name;
        }


}
