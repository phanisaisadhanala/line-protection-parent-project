package com.example.demo.entity;

import jakarta.persistence.*;
import java.time.LocalDateTime;

/**
 * Entity to store form submission data for audit and tracking purposes
 */
@Entity
@Table(name = "form_submissions")
public class FormSubmission {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    // Basic information
    @Column(name = "relay_location", length = 255)
    private String relayLocation;

    @Column(name = "line_number", length = 255)
    private String lineNumber;

    @Column(name = "remote_location", length = 255)
    private String remoteLocation;

    // System specifications
    @Column(name = "nominal_system_voltage")
    private String nominalSystemVoltage;

    @Column(name = "breaker_rating")
    private String breakerRating;

    @Column(name = "conductor_rating")
    private String conductorRating;

    // Complete form data as JSON
    @Lob
    @Column(name = "form_data_json", columnDefinition = "TEXT")
    private String formDataJson;

    // CSV file name (stored separately if needed)
    @Column(name = "csv_file_name")
    private String csvFileName;

    // Timestamp
    @Column(name = "uploaded_at", nullable = false)
    private LocalDateTime uploadedAt;

    @Column(name = "generated_file_name")
    private String generatedFileName;

    // Status tracking
    @Column(name = "status", length = 50)
    private String status; // e.g., "SUCCESS", "FAILED", "PROCESSING"

    @Column(name = "error_message", columnDefinition = "TEXT")
    private String errorMessage;

    // Constructors
    public FormSubmission() {
        this.uploadedAt = LocalDateTime.now();
        this.status = "PROCESSING";
    }

    public FormSubmission(String formDataJson) {
        this();
        this.formDataJson = formDataJson;
    }

    // Getters and Setters
    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getRelayLocation() {
        return relayLocation;
    }

    public void setRelayLocation(String relayLocation) {
        this.relayLocation = relayLocation;
    }

    public String getLineNumber() {
        return lineNumber;
    }

    public void setLineNumber(String lineNumber) {
        this.lineNumber = lineNumber;
    }

    public String getRemoteLocation() {
        return remoteLocation;
    }

    public void setRemoteLocation(String remoteLocation) {
        this.remoteLocation = remoteLocation;
    }

    public String getNominalSystemVoltage() {
        return nominalSystemVoltage;
    }

    public void setNominalSystemVoltage(String nominalSystemVoltage) {
        this.nominalSystemVoltage = nominalSystemVoltage;
    }

    public String getBreakerRating() {
        return breakerRating;
    }

    public void setBreakerRating(String breakerRating) {
        this.breakerRating = breakerRating;
    }

    public String getConductorRating() {
        return conductorRating;
    }

    public void setConductorRating(String conductorRating) {
        this.conductorRating = conductorRating;
    }

    public String getFormDataJson() {
        return formDataJson;
    }

    public void setFormDataJson(String formDataJson) {
        this.formDataJson = formDataJson;
    }

    public String getCsvFileName() {
        return csvFileName;
    }

    public void setCsvFileName(String csvFileName) {
        this.csvFileName = csvFileName;
    }

    public LocalDateTime getUploadedAt() {
        return uploadedAt;
    }

    public void setUploadedAt(LocalDateTime uploadedAt) {
        this.uploadedAt = uploadedAt;
    }

    public String getGeneratedFileName() {
        return generatedFileName;
    }

    public void setGeneratedFileName(String generatedFileName) {
        this.generatedFileName = generatedFileName;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public String getErrorMessage() {
        return errorMessage;
    }

    public void setErrorMessage(String errorMessage) {
        this.errorMessage = errorMessage;
    }

    @Override
    public String toString() {
        return "FormSubmission{" +
                "id=" + id +
                ", relayLocation='" + relayLocation + '\'' +
                ", lineNumber='" + lineNumber + '\'' +
                ", remoteLocation='" + remoteLocation + '\'' +
                ", uploadedAt=" + uploadedAt +
                ", status='" + status + '\'' +
                '}';
    }
}