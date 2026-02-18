package com.example.demo;

import com.example.demo.entity.FormSubmission;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

/**
 * Repository interface for FormSubmission entity
 * Provides database access methods for form submissions
 */
@Repository
public interface FormSubmissionRepository extends JpaRepository<FormSubmission, Long> {
    
    /**
     * Find submissions by relay location
     */
    List<FormSubmission> findByRelayLocationContainingIgnoreCase(String relayLocation);
    
    /**
     * Find submissions by line number
     */
    List<FormSubmission> findByLineNumberContainingIgnoreCase(String lineNumber);
    
    /**
     * Find submissions by status
     */
    List<FormSubmission> findByStatus(String status);
    
    /**
     * Find submissions between dates
     */
    List<FormSubmission> findByUploadedAtBetween(LocalDateTime start, LocalDateTime end);
    
    /**
     * Find the most recent submissions
     */
    List<FormSubmission> findTop10ByOrderByUploadedAtDesc();
    
    /**
     * Find by relay location and line number
     */
    Optional<FormSubmission> findByRelayLocationAndLineNumber(String relayLocation, String lineNumber);
    
    /**
     * Count submissions by status
     */
    @Query("SELECT COUNT(f) FROM FormSubmission f WHERE f.status = :status")
    long countByStatus(@Param("status") String status);
    
    /**
     * Delete old submissions (older than specified date)
     */
    void deleteByUploadedAtBefore(LocalDateTime date);
}