import javax.naming.Context;
import javax.naming.NamingEnumeration;
import javax.naming.NamingException;
import javax.naming.directory.Attributes;
import javax.naming.directory.SearchControls;
import javax.naming.directory.SearchResult;
import javax.naming.ldap.InitialLdapContext;
import javax.naming.ldap.LdapContext;
import java.io.IOException;
import java.util.*;

public class ActiveDirectory {

   private ArrayList<Employee>  employees = new ArrayList<Employee>();
   private String name;
   private String department;
   private String title;
   private String email;
   private  String number;
   private  String mobile;

    public void connectionToAD(String baseName, String sheetName) throws NamingException, IOException {


        int count =0;

        LdapContext ctx = null;
        try {
            Hashtable env = new Hashtable();
            String adminName = "CN=User\\, Test,OU=Shushary,OU=Accounts,OU=SPB,OU=Europe,OU=Cosma,OU=Magna Group,DC=magna,DC=global"; // пользователь у которого есть права для каких-либо "конкретных действий"..
            String adminPassword = "Cosma123";
            env.put(Context.INITIAL_CONTEXT_FACTORY, "com.sun.jndi.ldap.LdapCtxFactory");
            env.put(Context.SECURITY_AUTHENTICATION, "simple");
            //CN=User\, Test,OU=Shushary,OU=Accounts,OU=SPB,OU=Europe,OU=Cosma,OU=Magna Group,DC=magna,DC=global
            env.put(Context.SECURITY_PRINCIPAL, adminName);
            env.put(Context.SECURITY_CREDENTIALS, adminPassword);
            env.put(Context.PROVIDER_URL, "ldap://spbmsdc10.magna.global");//ldap :389, ldaps 636
           // System.out.println("Attempting to Connect...");

            ctx = new InitialLdapContext(env, null);
           // System.out.println("Connection Successful.");
        } catch (NamingException NEx) {
            System.out.println("LDAP Connection: FAILED");
            NEx.printStackTrace();
        }
        try {
            SearchControls searchCtls = new SearchControls();
            searchCtls.setSearchScope(SearchControls.SUBTREE_SCOPE);
            String searchFilter = "(&(objectClass=user)(CN=*))";
            //  ознакомиться с атрибутами, если null то все атрибуты
            //String returnedAtts[]={"memberOf", "sAMAccountName"}; // bla bla bla
            String returnedAtts[]= {"name" ,"department", "title", "mail", "telephoneNumber","mobile" };
            searchCtls.setReturningAttributes(returnedAtts);
            NamingEnumeration answer = ctx.search(baseName, searchFilter, searchCtls);

            while (answer.hasMoreElements()) {

                SearchResult sr = (SearchResult) answer.next();
                Attributes attrs = sr.getAttributes();

                if (attrs != null){
                    try {
                        if (attrs.get("mobile") != null) {
                            mobile = attrs.get("mobile").get().toString();
                        } else {
                            mobile = "empty";
                        }
                        if (attrs.get("mail") != null) {
                            email = attrs.get("mail").get().toString();
                        } else {
                            email = "empty";
                        }
                        if (attrs.get("telephoneNumber") != null) {
                            number = attrs.get("telephoneNumber").get().toString();
                        } else {
                            number = "empty";
                        }
                        if (attrs.get("name") != null) {
                            name = attrs.get("name").get().toString();
                        } else {
                            name =  "empty";
                        }
                        if (attrs.get("department") != null) {
                            department = attrs.get("department").get().toString();
                        } else {
                            department = "empty";
                        }
                        if (attrs.get("title") != null) {
                            title = attrs.get("title").get().toString();
                        } else {
                            title = "empty";
                        }

                        employees.add(new Employee(name, department, title, email, number, mobile));

                    } catch (Exception e) {
                        System.out.println("error");
                    }
                }
            }
            ctx.close();
        } catch (NamingException e) {
            System.err.println("Problem searching directory: " + e);
        }
        Comparator<Employee> comparable = Comparator.comparing(obj ->obj.getName());
        Collections.sort(employees, comparable);

        ExcelWriter excelWriter = new ExcelWriter();
        excelWriter.ExcelWrite(sheetName, employees);
    }



}
