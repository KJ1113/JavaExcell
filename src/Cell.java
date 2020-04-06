public class Cell implements Comparable<Cell>{
    private String name;
    private double ratio;
    public Cell(String name , double ratio){
        this.name = name;
        this.ratio =ratio;
    }
    public double getRatio() {
        return ratio;
    }
    public String getname() {
        return name;
    }
    @Override
    public int compareTo(Cell cell) {
        if(this.ratio < cell.ratio){
            return 1;
        }else if(this.ratio > cell.ratio){
            return -1;
        }else {
            return 0;
        }
    }
}
