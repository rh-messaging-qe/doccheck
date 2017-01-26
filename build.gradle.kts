import com.github.jengelman.gradle.plugins.shadow.ShadowPlugin
import com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar
import org.codehaus.groovy.tools.javac.JavaStubCompilationUnit
import org.gradle.api.tasks.Sync
import org.gradle.script.lang.kotlin.*
import org.jetbrains.kotlin.gradle.plugin.KotlinPluginWrapper
import org.junit.platform.gradle.plugin.JUnitPlatformExtension
import org.junit.platform.gradle.plugin.JUnitPlatformPlugin
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.io.File
import org.gradle.api.JavaVersion.VERSION_1_8

apply<ApplicationPlugin>()
apply<KotlinPluginWrapper>()
apply<ShadowPlugin>()

apply<JUnitPlatformPlugin>()

val project_version = DateTimeFormatter.BASIC_ISO_DATE.format(LocalDateTime.now())

val kotlin_version = "1.0.6"
val soffice_version = properties.getOrElse("oo.version", { "5.2.0" })
buildscript {
    val kotlin_version = "1.0.6"
    repositories {
        jcenter()
        mavenCentral()
        maven {
            setUrl("https://plugins.gradle.org/m2/")
        }
    }
    dependencies {
        classpath("org.jetbrains.kotlin:kotlin-gradle-plugin:$kotlin_version")
        classpath("org.junit.platform:junit-platform-gradle-plugin:1.0.0-M3")
        classpath("com.github.jengelman.gradle.plugins:shadow:1.2.3")
    }
}

configure<ApplicationPluginConvention> {
    mainClassName = "com.redhat.mqe.jms.main.TestRunner"
}

configure<JavaPluginConvention> {
    sourceCompatibility = VERSION_1_8
    targetCompatibility = VERSION_1_8
}

configure<JUnitPlatformExtension> {
    platformVersion = "1.0.0-M3"
}

// https://github.com/gradle/gradle-script-kotlin/blob/master/build.gradle.kts
tasks.withType<ShadowJar> {
    baseName = "msgqe-amq-docs"
    version = project_version
}

/**
 * Task installMacro overwrites old version of macro with new one
 *
 * @param -Poo.scriptsdir= e.g. /home/jdanek/.config/libreofficedev/4/user/Scripts
 */
task("installMacro", Sync::class) {
    val scriptsDir = project.findProperty("oo.scriptsdir") as String?
    if (scriptsDir != null) {
        val macroDir = "java/ComparisonModule/"
        from(tasks.withType<ShadowJar>().first().outputs)
        from(project.projectDir.resolve("src/main/parcel-descriptor.xml")) {
            expand(mutableMapOf<String, String>(
                    "projectVersion" to project_version
            ))
        }
        into(File(scriptsDir).resolve(macroDir).toString())
        filteringCharset = "UTF-8"
    }
}

repositories {
    jcenter()
    mavenCentral()
}

dependencies {
    compile("org.jetbrains.kotlin:kotlin-stdlib:$kotlin_version")

    compileClasspath("org.libreoffice:ridl:$soffice_version")
    compileClasspath("org.libreoffice:unoil:$soffice_version")
    testCompile("org.libreoffice:ridl:$soffice_version")
    testCompile("org.libreoffice:unoil:$soffice_version")
    testCompile("org.libreoffice:juh:$soffice_version")
    testCompile("org.libreoffice:jurt:$soffice_version")

    // https://github.com/junit-team/junit5-samples/tree/r5.0.0-M1/junit5-gradle-consumer
    testCompile("org.junit.jupiter:junit-jupiter-api:5.0.0-M3")
    testRuntime("org.junit.jupiter:junit-jupiter-engine:5.0.0-M3")

    testCompile(group = "com.google.truth", name = "truth", version = "0.30")
}
