using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using Google.Cloud.Firestore;
using OfficeOpenXml;

namespace ComparisonOfOrderPickingAlgorithms
{
    public class Program
    {
        private static Coordinate depot;
        private static Problem room;
        private static Parameters parameters;
        private static Picker picker;
        private static Solution solution;
        private static FirestoreDb database;
        private static Dictionary<string, object> dic = new Dictionary<string, object>();
        private static string filePath;


        //Method to make a set of item pick list tryouts to tune parameters of Genetic algorithm
        public static void tuneGeneticAlgorithmParameters(String pickListsFilePath, String outputFilePath, ExcelWorksheet DMworksheet)
        {
            String delimiter = "\t";
            StreamWriter wr = new StreamWriter(outputFilePath, true);
            wr.WriteLine("{0}" + delimiter + "{1}",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff"),
                pickListsFilePath);
            wr.WriteLine("instanceNumber" + delimiter + "DistanceMatrixRunningTime" + delimiter + "NumberOfStagnantGeneration" + delimiter + "PopulationSize" + delimiter + "CrossoverProbability" + delimiter + "MutationProbability" + delimiter + "CrossoverOperator" + delimiter + "MutationOperator" + delimiter + "TravelledDistance" + delimiter + "RunningTime");
            //wr.WriteLine("instanceNumber" + delimiter + "DistanceMatrixRunningTime" + delimiter + "NumberOfIterations" + delimiter + "PopulationSize" + delimiter + "CrossoverProbability" + delimiter + "MutationProbability" + delimiter + "CrossoverOperator" + delimiter + "MutationOperator" + delimiter + "TravelledDistance" + delimiter + "RunningTime");
            wr.Close();

            parameters.ItemListSet = Utils.readTestList(pickListsFilePath);

            bool distanceMatrixShouldBeCalculated = true;
            double[,] calculatedDistanceMatrix = new double[1, 1]; ;
            List<Coordinate>[,] calculatedPathMatrix = new List<Coordinate>[1, 1];
            double calculatedDistanceMatrixRunningTime = 0;

            //Add one additional Distance Matrix calculation at the beginning to initiate multi-core process and having less values for distance matrix calculation at report
            room.ItemList = Utils.Clone<Item>(parameters.ItemListSet.ElementAt(0));
            picker = new Picker(depot);
            solution = new Solution(room, picker, parameters);
            //solution.prepareDistanceMatrix(new Item(0, solution.Problem.NumberOfCrossAisles - 1, 1, 0, solution.Problem.S));
            //solution.importDistanceMatrix("../../../files/distance_matrix_by_floor.xlsx", new Item(0, room.NumberOfCrossAisles - 1, 1, 0, solution.Problem.S));


            for (int k = 0; k < parameters.ItemListSet.Count; k++)
            {
                List<Item> itemList = parameters.ItemListSet.ElementAt(k);
                for (int j = 0; j < parameters.NumberOfStagnantGenerationList.Length; j++)
                {
                    parameters.NumberOfStagnantGeneration = parameters.NumberOfStagnantGenerationList[j];
                    //for (int j = 0; j < parameters.NumberOfIterationsList.Length; j++)
                    //{
                    //parameters.NumberOfIterations = parameters.NumberOfIterationsList[j];
                    for (int i = 0; i < parameters.PopulationSizeList.Length; i++)
                    {
                        parameters.PopulationSize = parameters.PopulationSizeList[i];
                        for (int m = 0; m < parameters.CrossoverProbabilityList.Length; m++)
                        {
                            parameters.CrossoverProbability = parameters.CrossoverProbabilityList[m];
                            for (int n = 0; n < parameters.MutationProbabilityList.Length; n++)
                            {
                                parameters.MutationProbability = parameters.MutationProbabilityList[n];
                                for (int p = 0; p < parameters.CrossoverOperatorList.Length; p++)
                                {
                                    parameters.CrossoverOperator = parameters.CrossoverOperatorList[p];
                                    for (int r = 0; r < parameters.MutationOperatorList.Length; r++)
                                    {
                                        parameters.MutationOperator = parameters.MutationOperatorList[r];



                                        //Console.WriteLine("Solving #{0} pick list of test file: {1}", k + 1, pickListsFilePath);
                                        int trial = 5;
                                        double totalTD = 0;
                                        for (int ii = 0; ii < trial; ii++)
                                        {
                                            room.ItemList = Utils.Clone<Item>(itemList);
                                            picker = new Picker(depot);
                                            solution = new Solution(room, picker, parameters);
                                            if (distanceMatrixShouldBeCalculated)
                                            {
                                                //solution.prepareDistanceMatrix(new Item(0, solution.Problem.NumberOfCrossAisles - 1, 1, 0, solution.Problem.S));
                                                solution.prepareDistanceMatrixForSpecificPickList(room.ItemList, DMworksheet);
                                                calculatedDistanceMatrix = solution.DistanceMatrix;
                                                //calculatedPathMatrix = solution.PathMatrix;
                                                calculatedDistanceMatrixRunningTime = 0;
                                                distanceMatrixShouldBeCalculated = false;
                                            }
                                            else
                                            {
                                                //carrying already calculated distance matrix and path matrix to the solution
                                                solution.DistanceMatrix = calculatedDistanceMatrix;
                                                //solution.PathMatrix = calculatedPathMatrix;
                                                solution.DistanceMatrixRunningTime = calculatedDistanceMatrixRunningTime;
                                            }
                                            solution.solve(Solution.Algorithm.GeneticAlgorithm);
                                            totalTD += solution.TravelledDistance;
                                        }

                                        double avg_total_distance = totalTD / trial;

                                        Console.ForegroundColor = ConsoleColor.DarkGreen;
                                        Console.WriteLine("{0}" + delimiter + "{1}" + delimiter + "{2}" + delimiter + "{3}" + delimiter + "{4}" + delimiter + "{5}" + delimiter + "{6}" + delimiter + "{7}" + delimiter + "{8}" + delimiter + "{9}",
                                            k + 1,
                                            solution.DistanceMatrixRunningTime,
                                            parameters.NumberOfStagnantGeneration,
                                            //parameters.NumberOfIterations,
                                            parameters.PopulationSize,
                                            parameters.CrossoverProbability,
                                            parameters.MutationProbability,
                                            parameters.CrossoverOperator.ToString(),
                                            parameters.MutationOperator.ToString(),
                                            avg_total_distance,
                                            solution.RunningTime);
                                        //Console.WriteLine();
                                        Console.ResetColor();
                                        wr = new StreamWriter(outputFilePath, true);
                                        wr.WriteLine("{0}" + delimiter + "{1}" + delimiter + "{2}" + delimiter + "{3}" + delimiter + "{4}" + delimiter + "{5}" + delimiter + "{6}" + delimiter + "{7}" + delimiter + "{8}" + delimiter + "{9}",
                                            k + 1,
                                            solution.DistanceMatrixRunningTime,
                                            parameters.NumberOfStagnantGeneration,
                                            //parameters.NumberOfIterations,
                                            parameters.PopulationSize,
                                            parameters.CrossoverProbability,
                                            parameters.MutationProbability,
                                            parameters.CrossoverOperator.ToString(),
                                            parameters.MutationOperator.ToString(),
                                            avg_total_distance,
                                            solution.RunningTime);
                                        wr.Close();
                                    }
                                }
                            }
                        }
                    }
                }
                wr = new StreamWriter(outputFilePath, true);
                wr.Close();
                calculatedDistanceMatrix = null;
                calculatedPathMatrix = null;
                calculatedDistanceMatrixRunningTime = 0;
                distanceMatrixShouldBeCalculated = true;
            }
        }

        //Method to run our real world case with different algorithms
        public static void runRealWorldChallenge()
        {
            int S = 99;
            double W = 2.77;
            double L = 30.69;
            double K = 0.31;
            int no_of_horizontal_aisles = 2;
            int no_of_vertical_aisles = 39;

            depot = new Coordinate(1, no_of_horizontal_aisles);
            room = new Problem(S, W, L, K, no_of_horizontal_aisles - 1, no_of_vertical_aisles, depot);
            parameters = new Parameters();

            //parameters.NumberOfIterations = 100;
            parameters.NumberOfStagnantGeneration = 140;
            parameters.PopulationSize = 200;
            parameters.CrossoverProbability = 0.9f;
            parameters.MutationProbability = 0.1f;
            parameters.CrossoverOperator = PickListGAParameters.Crossover.PMX;
            parameters.MutationOperator = PickListGAParameters.Mutation.Inversion;
            picker = new Picker(depot);
            //Utils.generateTestLists(room, new int[] { 50 }, 5);
            parameters.ItemListSet = Utils.readTestList("../../../files/testListWithPickListSize005.txt");
            room.ItemList = Utils.Clone<Item>(parameters.ItemListSet.ElementAt(0));


            solution = new Solution(room, picker, parameters);


            //solution.prepareDistanceMatrix(new Item(0, solution.Problem.NumberOfCrossAisles - 1, 1, 0, solution.Problem.S));
            solution.importDistanceMatrix("../../../files/distance_matrix_by_floor.xlsx", new Item(0, room.NumberOfCrossAisles - 1, 1, 0, S));


            Console.WriteLine("----------S-Shape----------");
            solution.solve(Solution.Algorithm.SShape);
            WritePickListToDicForHeuristic(Solution.Algorithm.SShape, solution.Picker);
            Console.WriteLine("S-Shape TTD: {0}", solution.TravelledDistance);
            solution.Picker = new Picker(depot);
            room.ItemList = Utils.Clone<Item>(parameters.ItemListSet.ElementAt(0));

            Console.WriteLine("----------Largest Gap----------");
            solution.solve(Solution.Algorithm.LargestGap);
            WritePickListToDicForHeuristic(Solution.Algorithm.LargestGap, solution.Picker);
            Console.WriteLine("Largest Gap TTD: {0}", solution.TravelledDistance);
            solution.Picker = new Picker(depot);
            room.ItemList = Utils.Clone<Item>(parameters.ItemListSet.ElementAt(0));

            Console.WriteLine("----------MidPoint----------");
            solution.solve(Solution.Algorithm.MidPoint);
            WritePickListToDicForHeuristic(Solution.Algorithm.MidPoint, solution.Picker);
            Console.WriteLine("Mid Point TTD: {0}", solution.TravelledDistance);

            Console.WriteLine("----------Genetic----------");
            int nTrial = 5;
            double totalTD = 0;
            double minDistance = Int64.MaxValue;
            int[] depotFirstBestSolution = new int[0];
            Solution bestSolution = solution;
            for (int i = 0; i < nTrial; i++)
            {
                room = new Problem(S, W, L, K, no_of_horizontal_aisles - 1, no_of_vertical_aisles, depot);
                room.ItemList = Utils.Clone<Item>(parameters.ItemListSet.ElementAt(0));
                Solution solution_new = new Solution(room, new Picker(depot), parameters);
                solution_new.DistanceMatrix = solution.DistanceMatrix;
                solution_new.PathMatrix = solution.PathMatrix;
                int[] orderedSelectedItems = solution_new.solve(Solution.Algorithm.GeneticAlgorithm);
                if (solution_new.TravelledDistance < minDistance)
                {
                    minDistance = solution_new.TravelledDistance;
                    depotFirstBestSolution = orderedSelectedItems;
                    bestSolution = solution_new;
                }
                totalTD += solution_new.TravelledDistance;
                Console.WriteLine("Trial {0}: {1}", i, solution_new.TravelledDistance);
            }
            Console.WriteLine("Average Genetic TTD: {0}", totalTD / nTrial);

            extractPathForGenetic(depotFirstBestSolution, bestSolution);

            database.Collection("picklists").AddAsync(dic);
        }

        //Method to setup parameter tuning of Genetic Algorithm
        public static void setupGeneticAlgorithmParameterTuning(bool generateNewTestLists)
        {
            int S = 99;
            double W = 2.77;
            double L = 30.69;
            double K = 0.31;
            int no_of_horizontal_aisles = 2;
            int no_of_vertical_aisles = 39;

            depot = new Coordinate(1, no_of_horizontal_aisles);
            room = new Problem(S, W, L, K, no_of_horizontal_aisles - 1, no_of_vertical_aisles, depot);
            parameters = new Parameters();

            if (generateNewTestLists)
            {
                //Setup test list generation parameters here
                parameters.PickListSizesOfTestLists = new int[] { 20, 50, 100 };
                parameters.NumberOfPickLists = 5;
                //Utils.generateTestLists(room, parameters.PickListSizesOfTestLists, parameters.NumberOfPickLists);
            }

            String[] sizeNames = new String[] { "small", "medium", "large" };
            String[] itemListFilePaths = new String[]
            {
                "../../../files/testListWithPickListSize020.txt",
                "../../../files/testListWithPickListSize050.txt",
                "../../../files/testListWithPickListSize100.txt"
            };
            string distanceMatrixFilePath = "../../../files/distance_matrix_by_floor.xlsx";

            try
            {
                Console.WriteLine("Importing Distance Matrix. Please Wait...");
                FileInfo DMFile = new FileInfo(distanceMatrixFilePath);
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                using (var DMpackage = new ExcelPackage(DMFile))
                {
                    ExcelWorksheet DMworksheet = DMpackage.Workbook.Worksheets[0];

                    for (int i = 0; i < sizeNames.Length; i++)
                    {
                        if (i == 0)
                        {
                            parameters.NumberOfStagnantGenerationList = new int[] { 150 };
                            parameters.PopulationSizeList = new int[] { 200 };
                            parameters.CrossoverProbabilityList = new float[] { 0.6f };
                            parameters.MutationProbabilityList = new float[] { 0.1f };
                        } else if (i == 1)
                        {
                            parameters.NumberOfStagnantGenerationList = new int[] { 150 };
                            parameters.PopulationSizeList = new int[] { 200 };
                            parameters.CrossoverProbabilityList = new float[] { 0.9f };
                            parameters.MutationProbabilityList = new float[] { 0.05f };
                        } else
                        {
                            parameters.NumberOfStagnantGenerationList = new int[] { 150 };
                            parameters.PopulationSizeList = new int[] { 200 };
                            parameters.CrossoverProbabilityList = new float[] { 0.9f };
                            parameters.MutationProbabilityList = new float[] { 0.1f };
                        }
                        Console.WriteLine($"Parameter Tuning for {sizeNames[i]} picklist");
                        //parameters.NumberOfStagnantGenerationList = new int[] { 50, 150 };
                        ////parameters.NumberOfIterationsList = new int[] { 500, 1000, 2000 };
                        //parameters.PopulationSizeList = new int[] { 50, 100, 200 };
                        //parameters.CrossoverProbabilityList = new float[] { 0.6f, 0.9f };
                        //parameters.MutationProbabilityList = new float[] { 0.005f, 0.05f, 0.1f };
                        //parameters.CrossoverOperatorList = new PickListGAParameters.Crossover[] { PickListGAParameters.Crossover.PMX };
                        //parameters.MutationOperatorList = new PickListGAParameters.Mutation[] { PickListGAParameters.Mutation.Swap };
                        //parameters.NumberOfStagnantGenerationList = new int[] { 140 };
                        //parameters.PopulationSizeList = new int[] { 200 };
                        //parameters.CrossoverProbabilityList = new float[] { 0.9f };
                        //parameters.MutationProbabilityList = new float[] { 0.1f };
                        parameters.CrossoverOperatorList = new PickListGAParameters.Crossover[] { PickListGAParameters.Crossover.OX2, PickListGAParameters.Crossover.PMX, PickListGAParameters.Crossover.PositionBased };
                        parameters.MutationOperatorList = new PickListGAParameters.Mutation[] { PickListGAParameters.Mutation.Displacement, PickListGAParameters.Mutation.Insertion, PickListGAParameters.Mutation.Inversion, PickListGAParameters.Mutation.Shuffle, PickListGAParameters.Mutation.Swap };


                        String parameterTuningReportFilePath = $"../../../files/parameter_tuning/{sizeNames[i]}_operators.txt";

                        String path = itemListFilePaths[i];

                        if (File.Exists(path))
                        {
                            tuneGeneticAlgorithmParameters(path, parameterTuningReportFilePath, DMworksheet);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error on importing DM");
            }
        }

        public static void extractPathForGenetic(int[] depotFirstBestSolution, Solution bestSolution)
        {
            List<Coordinate> path = new List<Coordinate>();
            Item depotAsItem = new Item(0, bestSolution.Problem.NumberOfCrossAisles - 1, 1, 0, bestSolution.Problem.S);
            List<Coordinate> initialPath = getShorteshPathBetweenTwoItem(depotAsItem, bestSolution.IndexedItemDictionary[depotFirstBestSolution[1]], bestSolution.Problem.S);
            path.AddRange(initialPath);
            for (int i = 1; i < depotFirstBestSolution.Length - 1; i++)
            {
                int sourceIndex = depotFirstBestSolution[i];
                int targetIndex = depotFirstBestSolution[i + 1];
                List<Coordinate> p = getShorteshPathBetweenTwoItem(bestSolution.IndexedItemDictionary[sourceIndex], bestSolution.IndexedItemDictionary[targetIndex], bestSolution.Problem.S);
                path.AddRange(p);
            }
            List<Coordinate> finalPath = getShorteshPathBetweenTwoItem(bestSolution.IndexedItemDictionary[depotFirstBestSolution[depotFirstBestSolution.Length - 1]], depotAsItem, bestSolution.Problem.S);
            path.AddRange(finalPath);



            WriteItemsInOrderToDic(depotFirstBestSolution, bestSolution);
            WriteCoordinatesToDic(path);
            dic["geneticDistance"] = bestSolution.TravelledDistance;
        }

        static void WritePickListToDicForHeuristic(Solution.Algorithm algorithm, Picker picker)
        {
            string algorithmName = "";
            if (algorithm == Solution.Algorithm.SShape)
            {
                algorithmName = "sShape";
            }
            else if (algorithm == Solution.Algorithm.LargestGap)
            {
                algorithmName = "largestGap";
            }
            else if (algorithm == Solution.Algorithm.MidPoint)
            {
                algorithmName = "midPoint";
            }

            List<Dictionary<string, object>> items = new List<Dictionary<string, object>>();

            for (int i = 0; i < picker.PickedItems.Count; i++)
            {
                Item item = picker.PickedItems[i];
                if (item.Index == 0)
                {
                    continue;
                }
                Dictionary<string, object> itemDic = new Dictionary<string, object>()
                {
                    {"index", item.Index },
                    {"b", item.BInfo },
                    {"c", item.CInfo },
                    {"d", item.DInfo },
                };

                items.Add(itemDic);
            }

            List<Dictionary<string, object>> pathDicList = new List<Dictionary<string, object>>();

            List<Coordinate> path = picker.criticalPoints;

            for (int i = 0; i < path.Count; i++)
            {
                Coordinate coordinate = path[i];
                pathDicList.Add(new Dictionary<string, object>() { { "x", coordinate.X }, { "y", coordinate.Y }, });
            }

            dic[$"{algorithmName}Items"] = items;
            dic[$"{algorithmName}Path"] = pathDicList;
            dic[$"{algorithmName}Distance"] = picker.Distance;
        }

        static void WriteCoordinatesToDic(List<Coordinate> path)
        {
            List<Dictionary<string, object>> pathDicList = new List<Dictionary<string, object>>();
            for (int i = 0; i < path.Count; i++)
            {
                var coordinate = path[i];
                pathDicList.Add(new Dictionary<string, object>() { { "x", coordinate.X }, { "y", coordinate.Y }, });
            }

            dic["geneticPath"] = pathDicList;
        }

        static void WriteItemsInOrderToDic(int[] depotFirstBestSolution, Solution bestSolution)
        {
            List<Dictionary<string, object>> items = new List<Dictionary<string, object>>();
            for (int i = 1; i < depotFirstBestSolution.Length; i++)
            {
                Item item = bestSolution.IndexedItemDictionary[depotFirstBestSolution[i]];
                Dictionary<string, object> itemDic = new Dictionary<string, object>()
                {
                    {"index", item.Index },
                    {"b", item.BInfo },
                    {"c", item.CInfo },
                    {"d", item.DInfo },
                };

                items.Add(itemDic);
            }

            dic["geneticItems"] = items;
        }

        public static List<Coordinate> getShorteshPathBetweenTwoItem(Item source, Item target, int S)
        {
            int x1 = source.BInfo + source.CInfo;
            int x2 = target.BInfo + target.CInfo;
            int y1 = source.DInfo;
            int y2 = target.DInfo;
            List<Coordinate> myPath = new List<Coordinate>();

            myPath.Add(new Coordinate(x1, y1));
            if (x1 == x2)
            {
                myPath.Add(new Coordinate(x2, y2));
            }
            else if ((y1 + y2) <= S)
            {
                myPath.Add(new Coordinate(x1, 0));
                myPath.Add(new Coordinate(x2, 0));
                myPath.Add(new Coordinate(x2, y2));
            }
            else
            {
                myPath.Add(new Coordinate(x1, S));
                myPath.Add(new Coordinate(x2, S));
                myPath.Add(new Coordinate(x2, y2));
            }

            return myPath;
        }

        public static void initializeGeneticParameters(int picklistSize)
        {
            parameters = new Parameters();
            if (picklistSize <= 35)
            {
                parameters.NumberOfStagnantGeneration = 150;
                parameters.PopulationSize = 200;
                parameters.CrossoverProbability = 0.6f;
                parameters.MutationProbability = 0.1f;
                parameters.CrossoverOperator = PickListGAParameters.Crossover.PositionBased;
                parameters.MutationOperator = PickListGAParameters.Mutation.Inversion;
            }
            else if (picklistSize <= 75)
            {
                parameters.NumberOfStagnantGeneration = 150;
                parameters.PopulationSize = 200;
                parameters.CrossoverProbability = 0.9f;
                parameters.MutationProbability = 0.05f;
                parameters.CrossoverOperator = PickListGAParameters.Crossover.PMX;
                parameters.MutationOperator = PickListGAParameters.Mutation.Inversion;
            }
            else
            {
                parameters.NumberOfStagnantGeneration = 150;
                parameters.PopulationSize = 200;
                parameters.CrossoverProbability = 0.9f;
                parameters.MutationProbability = 0.1f;
                parameters.CrossoverOperator = PickListGAParameters.Crossover.PMX;
                parameters.MutationOperator = PickListGAParameters.Mutation.Inversion;
            }
        }

        public static void initializeGeneticParametersForTesting()
        {
            parameters = new Parameters();
            parameters.NumberOfStagnantGeneration = 150;
            parameters.PopulationSize = 4;
            parameters.CrossoverProbability = 0.6f;
            parameters.MutationProbability = 0.1f;
            parameters.CrossoverOperator = PickListGAParameters.Crossover.Cycle;
            parameters.MutationOperator = PickListGAParameters.Mutation.Swap;
        }

        public static void myApp()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "./picklists/";
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;

            int S = 99;
            double W = 2.77;
            double L = 30.69;
            double K = 0.31;
            int no_of_horizontal_aisles = 2;
            int no_of_vertical_aisles = 39;

            depot = new Coordinate(1, no_of_horizontal_aisles);
            room = new Problem(S, W, L, K, no_of_horizontal_aisles - 1, no_of_vertical_aisles, depot);
            picker = new Picker(depot);

            string distanceMatrixFilePath = "../../../files/distance_matrix_by_floor.xlsx";
            //string distanceMatrixFilePath = "./files/distance_matrix_by_floor.xlsx";
            try
            {
                Console.WriteLine("Importing Distance Matrix. Please Wait...");
                FileInfo DMFile = new FileInfo(distanceMatrixFilePath);
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                using (var DMpackage = new ExcelPackage(DMFile))
                {
                    ExcelWorksheet DMworksheet = DMpackage.Workbook.Worksheets[0];

                    string cellsPath = "../../../files/mezanin_cells_by_floors.xlsx";
                    //string cellsPath = "./files/mezanin_cells_by_floors.xlsx";
                    try
                    {
                        Console.WriteLine("Importing All Cells. Please Wait...");
                        FileInfo cellFile = new FileInfo(cellsPath);
                        ExcelPackage.LicenseContext = LicenseContext.Commercial;
                        using (var package = new ExcelPackage(cellFile))
                        {
                            ExcelWorksheet cellsWorksheet = package.Workbook.Worksheets[0];
                            while (true)
                            {
                                Console.WriteLine("Enter 1 to select file or enter file path");
                                string input = Console.ReadLine();

                                if (input == "1")
                                {
                                    DialogResult result = openFileDialog.ShowDialog();
                                    if (result != null)
                                    {
                                        filePath = openFileDialog.FileName;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                else
                                {
                                    filePath = input;
                                }


                                List<Item> items = ReadPickListFromExcelFile(filePath, cellsWorksheet);
                                initializeGeneticParameters(items.Count);
                                //initializeGeneticParametersForTesting();
                                room.ItemList = Utils.Clone<Item>(items);
                                solution = new Solution(room, picker, parameters);
                                solution.prepareDistanceMatrixForSpecificPickList(items, DMworksheet);
                                solveUsingAllAlgorithms(items);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error on DM import");
            }
        }

        public static void solveUsingAllAlgorithms(List<Item> items)
        {
            Solution solution_new;

            Console.WriteLine("----------S-Shape----------");
            solution_new = new Solution(room, new Picker(depot), parameters);
            solution_new.solve(Solution.Algorithm.SShape);
            WritePickListToDicForHeuristic(Solution.Algorithm.SShape, solution_new.Picker);
            Console.WriteLine("S-Shape TTD: {0}", solution_new.TravelledDistance);

            Console.WriteLine("----------Largest Gap----------");
            room.ItemList = Utils.Clone<Item>(items);
            solution_new = new Solution(room, new Picker(depot), parameters);
            solution_new.solve(Solution.Algorithm.LargestGap);
            WritePickListToDicForHeuristic(Solution.Algorithm.LargestGap, solution_new.Picker);
            Console.WriteLine("Largest Gap TTD: {0}", solution_new.TravelledDistance);


            Console.WriteLine("----------MidPoint----------");
            room.ItemList = Utils.Clone<Item>(items);
            solution_new = new Solution(room, new Picker(depot), parameters);
            solution_new.solve(Solution.Algorithm.MidPoint);
            WritePickListToDicForHeuristic(Solution.Algorithm.MidPoint, solution_new.Picker);
            Console.WriteLine("Mid Point TTD: {0}", solution_new.TravelledDistance);

            Console.WriteLine("----------Genetic----------");
            int nTrial = 5;
            double totalTD = 0;
            double minDistance = Int64.MaxValue;
            int[] depotFirstBestSolution = new int[0];
            Solution bestSolution = solution_new;
            for (int i = 0; i < nTrial; i++)
            {
                room.ItemList = Utils.Clone<Item>(items);
                Solution genetic_solution = new Solution(room, new Picker(depot), parameters);
                genetic_solution.DistanceMatrix = solution.DistanceMatrix;
                int[] orderedSelectedItems = genetic_solution.solve(Solution.Algorithm.GeneticAlgorithm);
                if (genetic_solution.TravelledDistance < minDistance)
                {
                    minDistance = genetic_solution.TravelledDistance;
                    depotFirstBestSolution = orderedSelectedItems;
                    bestSolution = genetic_solution;
                }
                totalTD += genetic_solution.TravelledDistance;
                Console.WriteLine("Trial {0}: {1}", i, genetic_solution.TravelledDistance);
            }
            Console.WriteLine("Average Genetic TTD: {0}", totalTD / nTrial);

            extractPathForGenetic(depotFirstBestSolution, bestSolution);

            database.Collection("picklists").Document(Path.GetFileNameWithoutExtension(filePath)).SetAsync(dic);
        }

        public static List<Item> ReadPickListFromExcelFile(string filePath, ExcelWorksheet cellsWorksheet)
        {
            List<Item> items = new List<Item>();
            try
            {
                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                using (var package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int columnCount = worksheet.Dimension.Rows;
                    for (int i = 2; i <= columnCount; i++)
                    {
                        string code = worksheet.Cells[i, 2].Value?.ToString();
                        items.Add(getItemInfosFromCells(code, i - 1, cellsWorksheet));
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("asdfsdf");
                Console.WriteLine($"An error occurred while importing the distance matrix: {ex.Message}");
            }
            return items;
        }

        public static Item getItemInfosFromCells(string code, int index, ExcelWorksheet cellsWorksheet)
        {
            int bInfo;
            int cInfo;
            int dInfo;
            int rowIndex = -1;
            int colIndex = -1;
            for (int row = cellsWorksheet.Dimension.Start.Row; row <= cellsWorksheet.Dimension.End.Row; row++)
            {
                for (int col = cellsWorksheet.Dimension.Start.Column; col <= cellsWorksheet.Dimension.End.Column; col++)
                {
                    object cellValue = cellsWorksheet.Cells[row, col].Value;
                    if (cellValue != null && cellValue.ToString() == code)
                    {
                        rowIndex = row;
                        colIndex = col;
                        break;
                    }
                }
                if (rowIndex != -1)
                {
                    break; // Exit the outer loop if the value is found
                }
            }
            if (rowIndex != -1 && colIndex != -1)
            {
                dInfo = 100 - colIndex;
                if (rowIndex <= 4)
                {
                    bInfo = 39;
                    cInfo = 0;
                }
                else
                {
                    bInfo = 40 - ((rowIndex - 6) / 10 + 2);
                    cInfo = ((rowIndex - 7) % 8) < 4 ? 1 : 0;
                }
                if (bInfo == 1)
                {
                    cInfo = 1;
                }
                return new Item(index, 1, bInfo, cInfo, dInfo);


            }
            else
            {
                throw new Exception("Code not found in cells workbook");
            }
        }

        public static int getBInfo(string code)
        {
            int zoneAdder = 0;
            switch (code[0])
            {
                case 'A':
                    zoneAdder = 1;
                    break;
                case 'B':
                    zoneAdder = 4;
                    break;
                case 'C':
                    zoneAdder = 13;
                    break;
            }

            return 40 - zoneAdder + (int.Parse(code.Substring(2, 2)) / 2);
        }

        public static int getCInfo(string code)
        {
            return (int.Parse(code.Substring(2, 2)) % 2) == 0 ? 1 : 0;
        }

        public static int getDInfo(string code)
        {
            return 100 - ((int.Parse(code.Substring(5, 2)) - 1) * 9 + int.Parse(code.Substring(8, 2)));
        }

        [STAThread]
        public static void Main(string[] args)
        {
            string filepath = "../../../warehouse-illustration-firebase-adminsdk-urhj7-d89d276479.json";
            //string filepath = "./files/warehouse-illustration-firebase-adminsdk-urhj7-d89d276479.json";
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", filepath);
            database = FirestoreDb.Create("warehouse-illustration");
            myApp();


            //runRealWorldChallenge();
            //setupGeneticAlgorithmParameterTuning(true);
            //Console.ReadLine();
        }
    }
}