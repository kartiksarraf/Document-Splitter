plugins {
    id("java")
}

group = "com.xebia.documentSplitter"
version = "1.1"

repositories {
    mavenCentral()
}

java {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
}

dependencies {
        implementation("org.apache.poi:poi:5.2.4") // latest version for poi
        implementation("org.apache.poi:poi-ooxml:5.4.0") {
            exclude(group= "org.apache.commons", module= "commons-compress")
            exclude(group= "org.apache.commons", module= "commons-io")
        }
        implementation("org.apache.poi:poi-scratchpad:5.2.1")

        implementation("org.slf4j:slf4j-api:2.0.7") // keeping SLF4J API same as there are no new versions
        implementation("org.slf4j:slf4j-simple:2.0.7") // keeping slf4j-simple version same as there are no new versions

        implementation("org.apache.logging.log4j:log4j-core:2.20.0") // latest version for log4j-core
        implementation("org.apache.logging.log4j:log4j-api:2.20.0") // latest version for log4j-api

        implementation("org.apache.commons:commons-collections4:4.4")

        implementation("org.apache.commons:commons-compress:1.26.0")
        implementation("commons-io:commons-io:2.18.0")

        implementation("org.apache.xmlbeans:xmlbeans:5.1.1") // no change, it's already up to date

        compileOnly(fileTree(mapOf("dir" to "libs", "include" to listOf("*.jar"))))

        testImplementation(platform("org.junit:junit-bom:5.10.0")) // latest JUnit BOM
        testImplementation("org.junit.jupiter:junit-jupiter:5.10.0") // latest JUnit Jupiter version

}

tasks.test {
    useJUnitPlatform()
}

tasks.named<Jar>("jar") {
    duplicatesStrategy = DuplicatesStrategy.FAIL
    into("META-INF/lib") {
        from(
        configurations.runtimeClasspath.get()
        .filter { file ->
        // List of dependencies to include
        file.name.contains("poi") ||
        file.name.contains("slf4j") ||
        file.name.contains("log4j") ||
        file.name.contains("commons-collections4") ||
        file.name.contains("commons-compress") ||
        file.name.contains("commons-io") ||
        file.name.contains("xmlbeans")
        }
        )
    }

    from(sourceSets.main.get().allSource) {
        into("src")
    }

    manifest {
        attributes["Spring-Context"] = "*;publish-context:=false"
    }
}