import javax.naming.NamingException;
import java.io.IOException;

public class ActiveMain {

    public static void main(String[] args) throws IOException, NamingException {


      /*  AD  ad = new AD();
        ad.connectionToAD("OU=Kamenka,OU=Accounts,OU=SPB,OU=Europe,OU=Cosma,OU=Magna Group,DC=magna,DC=global", "Kamenka");
        ad.connectionToAD("OU=Shushary,OU=Accounts,OU=SPB,OU=Europe,OU=Cosma,OU=Magna Group,DC=magna,DC=global", "Shushary");
            */
        ActiveDirectory activeDirectory = new ActiveDirectory();
        activeDirectory.connectionToAD("OU=Kamenka,OU=Accounts,OU=SPB,OU=Europe,OU=Cosma,OU=Magna Group,DC=magna,DC=global", "Kamenka");
        activeDirectory.connectionToAD("OU=Shushary,OU=Accounts,OU=SPB,OU=Europe,OU=Cosma,OU=Magna Group,DC=magna,DC=global", "Shushary");

    }
}
