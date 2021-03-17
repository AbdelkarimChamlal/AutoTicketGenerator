package models;

/**
 * Ticket class is used as data holder or a schema
 * to convert xlsx rows into it
 * each Ticket object should be linked to a row in the input file
 * the main rule of this class is to hold data
 * and it is only setters and getters
 * with two constructors methode
 * one which gives the value "" to all params in case of nothing provided
 * the other takes all the values to params as input and places them to their params.
 *
 * @param chain which contains the ticket chain
 * @param base which contains the ticket base
 * @param wireType which contains the ticket wireType
 * @param process which contains the ticket process
 * @param skNumber which contains the ticket skNumber
 * @param followUp which contains the ticket followUp
 * @param corA which contains the ticket corA
 * @param corB which contains the ticket corB
 * @param wireCrossSection which contains the ticket wireCrossSection
 * @param insertion which contains the ticket insertion
 * @param post which contains the ticket post
 * @param sequence which contains the ticket sequence
 * @param size which contains the ticket size
 */
public class Ticket {

    String chain,base,wireType,process,skNumber,followUp,corA,corB,wireCrossSection,insertion,post,sequence,size;


    public String getChain() {
        return chain;
    }

    public void setChain(String chain) {
        this.chain = chain;
    }

    public String getBase() {
        return base;
    }

    public void setBase(String base) {
        this.base = base;
    }

    public String getWireType() {
        return wireType;
    }

    public void setWireType(String wireType) {
        this.wireType = wireType;
    }

    public String getProcess() {
        return process;
    }

    public void setProcess(String process) {
        this.process = process;
    }

    public String getSkNumber() {
        return skNumber;
    }

    public void setSkNumber(String skNumber) {
        this.skNumber = skNumber;
    }

    public String getFollowUp() {
        return followUp;
    }

    public void setFollowUp(String followUp) {
        this.followUp = followUp;
    }

    public String getCorA() {
        return corA;
    }

    public void setCorA(String corA) {
        this.corA = corA;
    }

    public String getCorB() {
        return corB;
    }

    public void setCorB(String corB) {
        this.corB = corB;
    }

    public String getWireCrossSection() {
        return wireCrossSection;
    }

    public void setWireCrossSection(String wireCrossSection) {
        this.wireCrossSection = wireCrossSection;
    }

    public String getInsertion() {
        return insertion;
    }

    public void setInsertion(String insertion) {
        this.insertion = insertion;
    }

    public String getPost() {
        return post;
    }

    public void setPost(String post) {
        this.post = post;
    }

    public String getSequence() {
        return sequence;
    }

    public void setSequence(String sequence) {
        this.sequence = sequence;
    }

    public String getSize() {
        return size;
    }

    public void setSize(String size) {
        this.size = size;
    }

    public Ticket() {
        this.base = "";
        this.chain = "";
        this.wireType = "";
        this.process = "";
        this.skNumber = "";
        this.followUp = "";
        this.corA = "";
        this.corB = "";
        this.wireCrossSection = "";
        this.insertion = "";
        this.post = "";
        this.sequence = "";
        this.size = "";
    }

    @Override
    public String toString() {
        return "Ticket{" +
                "chain='" + chain + '\'' +
                ", base='" + base + '\'' +
                ", wireType='" + wireType + '\'' +
                ", process='" + process + '\'' +
                ", skNumber='" + skNumber + '\'' +
                ", followUp='" + followUp + '\'' +
                ", corA='" + corA + '\'' +
                ", corB='" + corB + '\'' +
                ", wireCrossSection='" + wireCrossSection + '\'' +
                ", insertion='" + insertion + '\'' +
                ", post='" + post + '\'' +
                ", sequence='" + sequence + '\'' +
                ", size='" + size + '\'' +
                '}';
    }

    public Ticket(String chain, String base, String wireType, String process, String skNumber, String followUp, String corA, String corB, String wireCrossSection, String insertion, String post, String sequence, String size) {
        this.chain = chain;
        this.base = base;
        this.wireType = wireType;
        this.process = process;
        this.skNumber = skNumber;
        this.followUp = followUp;
        this.corA = corA;
        this.corB = corB;
        this.wireCrossSection = wireCrossSection;
        this.insertion = insertion;
        this.post = post;
        this.sequence = sequence;
        this.size = size;
    }

}
