import java.util.ArrayList;
import java.util.Iterator;

public class IteratorTest {

    private static class People {
        public People(String name, String age) {
            this.name = name;
            this.age = age;
        }

        private String name;
        private String age;

        public void setName(String name) {
            this.name = name;
        }

        public void setAge(String age) {
            this.age = age;
        }

        @Override
        public String toString() {
            return "People{" +
                    "name='" + name + '\'' +
                    ", age='" + age + '\'' +
                    '}';
        }
    }

    public static void main(String[] args) {
        String s = null;
        System.out.println(s != null && s.isEmpty());
    }
}
