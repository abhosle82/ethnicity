package enablement.transformers.hackathon.ethnicity;

import enablement.transformers.hackathon.ethnicity.entities.EthnicityBean;
import enablement.transformers.hackathon.ethnicity.poi.FLNameBusinessConatct1;
import enablement.transformers.hackathon.ethnicity.poi.FLNamesBusinessContact2;
import enablement.transformers.hackathon.ethnicity.poi.ReadFirstNameLastName;
import enablement.transformers.hackathon.ethnicity.poi.WriteFirstNameLastName;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.List;

@SpringBootApplication
public class EthinicityApplication  implements CommandLineRunner {

    public static void main(String[] args) {

        SpringApplication.run(EthinicityApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {

        List<EthnicityBean> lsEB = FLNamesBusinessContact2.readDiversityAndInclusionData();
        WriteFirstNameLastName.writeDiversityAndInclusionData(lsEB);


    }
}
