# ---- Build Stage ----
FROM maven:3.9.6-eclipse-temurin-17 AS build

WORKDIR /app
COPY . .
RUN mvn clean package -DskipTests

# ---- Run Stage ----
FROM eclipse-temurin:17-jre-alpine

WORKDIR /app
COPY --from=build /app/target/line-protection-calculator-1.0.0.jar app.jar

EXPOSE 8080

CMD ["java", "-jar", "app.jar"]
