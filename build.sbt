name := "fancy-poi"

organization := "org.fancypoi"

version := "1.0"

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "3.7",
  "org.apache.poi" % "poi-ooxml" % "3.7"
)

scalaVersion := "2.10.0"
