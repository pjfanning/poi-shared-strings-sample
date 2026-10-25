# poi-shared-strings-sample
Sample demonstrating the use of [poi-shared-strings](https://github.com/pjfanning/poi-shared-strings) when reading or writing xlsx files.

Requires Java 17 or later.

`./gradlew clean run`

The build.gradle file has a mainClass with a default of `com.github.pjfanning.poi.sample.XLSX2CSV`.
This is the read sample.

Change the mainClass in the `application` block to `com.github.pjfanning.poi.sample.StreamingWrite` to run the write sample. 
