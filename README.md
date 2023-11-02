# DocumentLoader + JLOP Helper Classes

This project is derived from the [OpenOffice DocumentHandling example](https://wiki.openoffice.org/wiki/API/Samples/Java/Office/DocumentHandling).
Next to the original approach of the example the [Java LibreOffice Programming (JLOP)](https://flywire.github.io/lo-p/index.html) helper classes where added. Since by default they don't have a package, they were moved to the `helper` package.
For compile reasons the corresponding dependencies for LibreOffice 7.3 are part of this repository.

You can find the derived example in the class `org.openoffice.sdk.example.documenthandling.LoLoader`, [here](src/main/java/org/openoffice/sdk/example/documenthandling).

## How to build
Please make sure you have a valid `JAVA_HOME` environment variable.
You can download the gradle binaries [here](https://gradle.org/). In case you have no local gradle installation you can use the wrapper scripts provided by this project. On Max/Linux use `gradlew`, on Windows `gradle.bat`

This is a gradle project, you can simply run `gradle build`.


This will create a Fat Jar, containing all the necessary classes for the mofuzz implementation. You can find the JAR in `build/libs/LoHelper-Lo.jar`.
