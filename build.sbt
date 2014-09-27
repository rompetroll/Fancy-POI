name := "fancy-poi"

organization := "org.fancypoi"

version := "1.0"

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "3.10.1",
  "org.apache.poi" % "poi-ooxml" % "3.10.1"
)

scalaVersion := "2.10.4"
