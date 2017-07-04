package org.qcri.rheem.apps.sindy

import java.lang.Iterable
import java.util

import de.hpi.isg.profiledb.store.model.Experiment
import org.qcri.rheem.api._
import org.qcri.rheem.apps
import org.qcri.rheem.apps.sindy.Sindy.{CellCreator, CellMerger, IndCandidateGenerator, IndCandidateMerger}
import org.qcri.rheem.apps.util.{Parameters, ProfileDBHelper, StdOut}
import org.qcri.rheem.core.api.{Configuration, RheemContext}
import org.qcri.rheem.core.function.FunctionDescriptor.{SerializableBinaryOperator, SerializableFunction}
import org.qcri.rheem.core.optimizer.ProbabilisticDoubleInterval
import org.qcri.rheem.core.plugin.Plugin
import org.qcri.rheem.core.util.fs.FileSystems

/**
  * This is a Rheem-based implementation of the SINDY algorithm.
  */
class Sindy(plugins: Plugin*) {

  /**
    * Execute the SINDY algorithm.
    *
    * @param paths         input file or directory URLs
    * @param seperator     CSV separator in the files
    * @param configuration Rheem configuration
    * @param experiment    the experiment to log measurements to
    * @return the INDs
    */
  def apply(paths: Seq[String], seperator: Char = ';')
           (implicit configuration: Configuration, experiment: Experiment) = {
    val rheemContext = new RheemContext(configuration)
    plugins.foreach(rheemContext.register)
    val planBuilder = new PlanBuilder(rheemContext)
      .withJobName(s"Sindy ($paths)")
      .withExperiment(experiment)
      .withUdfJarsOf(classOf[Sindy])

    val fileColumnIdOffsets = paths.flatMap(resolveDirs).zipWithIndex.map { case (url, index) => (url, index * 1000) }
    val allCells = fileColumnIdOffsets
      .map { case (path, offset) =>
        planBuilder
          .readTextFile(path).withName(s"Load $path")
          .flatMapJava(new CellCreator(offset, seperator),
            udfSelectivity = ProbabilisticDoubleInterval.createFromSpecification(
              "my.udf.Sindy.flatmap1-" + paths, configuration
            ),
            udfSelectivityKey = "my.udf.Sindy.flatmap1-" + paths
          ).withName(s"my.udf.Sindy.flatmap1-Create cells for $path")
      }
      .reduce(_ union _)

    val rawInds = allCells
      .map(cell => (cell._1, Array(cell._2))).withName("Prepare cell merging")
      .reduceByKeyJava(toSerializableFunction(_._1), new CellMerger,
        udfSelectivity = ProbabilisticDoubleInterval.createFromSpecification(
          "my.udf.Sindy.reduceBy1-" + paths, configuration
        ),
        udfSelectivityKey = "my.udf.Sindy.reduceBy1-" + paths
      ).withName("my.udf.Sindy.reduceBy1-Merge cells")
      .flatMapJava(new IndCandidateGenerator,
        udfSelectivity = ProbabilisticDoubleInterval.createFromSpecification(
          "my.udf.Sindy.flatmap2-" + paths, configuration
        ),
        udfSelectivityKey = "my.udf.Sindy.flatmap2-" + paths
      ).withName("my.udf.Sindy.flatmap2-Generate IND candidate sets")
      .reduceByKeyJava(toSerializableFunction(_._1), new IndCandidateMerger,
        udfSelectivity = ProbabilisticDoubleInterval.createFromSpecification(
          "my.udf.Sindy.reduceBy2-" + paths, configuration
        ),
        udfSelectivityKey = "my.udf.Sindy.reduceBy2-" + paths
      ).withName("my.udf.Sindy.reduceBy2-Merge IND candidate sets")
      .filter(_._2.length > 0,
        udfSelectivity = ProbabilisticDoubleInterval.createFromSpecification(
          "my.udf.Sindy.filter1-" + paths, configuration
        ),
        udfSelectivityKey = "my.udf.Sindy.filter1-" + paths
      ).withName("my.udf.Sindy.filter1-Filter empty candidate sets")
      .collect()

    def resolveColumnId(id: Int) = fileColumnIdOffsets
      .find { case (file, offset) => offset <= id && offset + 1000 > id } match {
      case Some((file, offset)) => s"$file[${id - offset}]"
      case _ => s"???[$id]"
    }

    rawInds.map {
      case (dep, refs) => (s"${resolveColumnId(dep)}", refs.map(resolveColumnId).toSeq)
    }
  }

  /**
    * If the given URL is a directory, list all its files recursively.
    *
    * @param url the URL
    * @return the file URLs
    */
  def resolveDirs(url: String): scala.Iterable[String] = {
    import scala.collection.JavaConversions._
    val fs = FileSystems.requireFileSystem(url)
    if (fs.isDirectory(url)) fs.listChildren(url).flatMap(resolveDirs) else Seq(url)
  }

}

/**
  * Companion object for [[Sindy]].
  */
object Sindy extends apps.util.ExperimentDescriptor {

  def version = "0.1.0"

  def main(args: Array[String]): Unit = {
    // Parse parameters.
    if (args.isEmpty) {
      sys.error(s"Usage: <main class> ${Parameters.experimentHelp} <plugin>(,<plugin>)* <CSV separator> <input URL>(;<input URL>)*")
      sys.exit(1)
    }

    implicit val configuration = new Configuration
    implicit val experiment = Parameters.createExperiment(args(0), this)
    val plugins = Parameters.loadPlugins(args(1))
    experiment.getSubject.addConfiguration("plugins", args(1))
    val separator = if (args(2).length == 1) args(2).charAt(0) else args(2) match {
      case "tab" => '\t'
      case "\\t" => '\t'
      case "comma" => ','
      case "semicolon" => ';'
      case "\\|" => '|'
      case "pipe" => '|'
      case other: String => throw new IllegalArgumentException("Unknown separator.")
    }
    val inputUrls = args(3).split(";")
    experiment.getSubject.addConfiguration("inputs", inputUrls)

    // Prepare the PageRank.
    val sindy = new Sindy(plugins: _*)

    // Run the PageRank.
    val inds = sindy(inputUrls.toSeq, separator).toSeq

    // Store experiment data.
    val inputFileSizes = inputUrls.map(url => FileSystems.getFileSize(url))
    if (inputFileSizes.forall(_.isPresent))
      experiment.getSubject.addConfiguration("inputSize", inputFileSizes.map(_.getAsLong).sum)
    ProfileDBHelper.store(experiment, configuration)

    // Print the result.
    StdOut.printLimited(inds)
  }

  /**
    * UDF to parse a CSV row and create cells.
    *
    * @param offset the column ID offset for the input CSV rows
    */
  class CellCreator(val offset: Int, val separator: Char) extends SerializableFunction[String, java.lang.Iterable[(String, Int)]] {

    override def apply(row: String): Iterable[(String, Int)] = {
      val fields = row.split(separator)
      val cells = new util.ArrayList[(String, Int)](fields.length)
      var columnId = offset
      for (field <- fields) {
        cells.add((field, columnId))
        columnId += 1
      }
      cells
    }
  }

  /**
    * UDF to merge the column IDs of two cells.
    */
  class CellMerger extends SerializableBinaryOperator[(String, Array[Int])] {

    import scala.collection.mutable

    lazy val merger = mutable.Set[Int]()

    override def apply(cell1: (String, Array[Int]), cell2: (String, Array[Int])): (String, Array[Int]) = {
      merger.clear()
      for (columnId <- cell1._2) merger += columnId
      for (columnId <- cell2._2) merger += columnId
      (cell1._1, merger.toArray)
    }
  }

  /**
    * UDF to create IND candidates from a cell group.
    */
  class IndCandidateGenerator extends SerializableFunction[(String, Array[Int]), java.lang.Iterable[(Int, Array[Int])]] {

    override def apply(cellGroup: (String, Array[Int])): java.lang.Iterable[(Int, Array[Int])] = {
      val columnIds = cellGroup._2
      val result = new util.ArrayList[(Int, Array[Int])](columnIds.length)
      for (i <- columnIds.indices) {
        val refColumnIds = new Array[Int](columnIds.length - 1)
        java.lang.System.arraycopy(columnIds, 0, refColumnIds, 0, i)
        java.lang.System.arraycopy(columnIds, i + 1, refColumnIds, i, refColumnIds.length - i)
        result.add((columnIds(i), refColumnIds))
      }
      result
    }
  }

  /**
    * UDF to merge two IND candidates.
    */
  class IndCandidateMerger extends SerializableBinaryOperator[(Int, Array[Int])] {

    import scala.collection.mutable

    lazy val merger = mutable.Set[Int]()

    override def apply(indc1: (Int, Array[Int]), indc2: (Int, Array[Int])): (Int, Array[Int]) = {
      merger.clear()
      for (columnId <- indc1._2) merger += columnId
      (indc1._1, indc2._2.filter(merger.contains))
    }

  }


}
