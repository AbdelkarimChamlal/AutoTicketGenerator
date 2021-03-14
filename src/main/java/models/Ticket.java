package models;

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
