package example.repository;

import example.entity.ExcelExample;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ExcelExampleRepository extends JpaRepository<ExcelExample, Long> {

}
