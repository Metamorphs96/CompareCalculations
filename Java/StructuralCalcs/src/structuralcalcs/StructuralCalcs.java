
public class StructuralCalcs {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        double Cpe;
        double qz;
        double pn;
        double w;
        double L;
        double M;

        Cpe=-0.7;
        qz=0.96;			//kPa
        pn=Cpe*qz;			//kPa
        w=pn*3.000;  			//kN/m
        L=6.000;			//m
        M=w*Math.pow(L,2)/8;            //kNm   
        
        System.out.format("Moment %8.2f kNm%n", M);
    }
    
}
