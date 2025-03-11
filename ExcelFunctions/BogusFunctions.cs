using Bogus;
using ExcelDna.Integration;
using ExcelFunctions.Services;
using FuzzySharp;
using Markdig;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using System;
using System.Data.Common;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static Bogus.DataSets.Name;

namespace ExcelFunctions
{
    public static class BogusFunctions
    {
        // Parameters:
        //   gender:
        //     For locale's that support Gender naming.

        [ExcelFunction(Category = "String", Description = "Get a first name.", HelpTopic = "Get a first name.")]
        public static object BOGUSFIRSTNAME(
    [ExcelArgument("[gender]", Name = "[gender]", Description = "(Optional) For locale's that support Gender naming.")] object gender,
    [ExcelArgument("[locale]", Name = "[locale]", Description = "(Optional) The current locale for the dataset.")] object locale,
    [ExcelArgument("[seed]", Name = "[seed]", Description = "(Optional) A number used to calculate a starting value for the pseudo-random number sequence.")] object seed)
        {
            Faker faker = CreateFaker(locale, seed);

            return BogusGenerateWithGender(gender, locale, seed, faker, faker.Name.FirstName);
        }

        [ExcelFunction(Category = "String", Description = "Get a last name.", HelpTopic = "Get a last name.")]
        public static object BOGUSLASTNAME(
[ExcelArgument("[gender]", Name = "[gender]", Description = "(Optional) For locale's that support Gender naming.")] object gender,
[ExcelArgument("[locale]", Name = "[locale]", Description = "(Optional) The current locale for the dataset.")] object locale,
[ExcelArgument("[seed]", Name = "[seed]", Description = "(Optional) A number used to calculate a starting value for the pseudo-random number sequence.")] object seed)
        {
            Faker faker = CreateFaker(locale, seed);

            return BogusGenerateWithGender(gender, locale, seed, faker, faker.Name.LastName);
        }

        [ExcelFunction(Category = "String", Description = "Gets a random job title.", HelpTopic = "Gets a random job title.")]
        public static object BOGUSJOBTITLE(
[ExcelArgument("[locale]", Name = "[locale]", Description = "(Optional) The current locale for the dataset.")] object locale,
[ExcelArgument("[seed]", Name = "[seed]", Description = "(Optional) A number used to calculate a starting value for the pseudo-random number sequence.")] object seed)
        {
            Faker faker = CreateFaker(locale, seed);

            return faker.Name.JobTitle();
        }

        [ExcelFunction(Category = "String", Description = "Get a type of job.", HelpTopic = "Get a type of job.")]
        public static object BOGUSJOBTYPE(
[ExcelArgument("[locale]", Name = "[locale]", Description = "(Optional) The current locale for the dataset.")] object locale,
[ExcelArgument("[seed]", Name = "[seed]", Description = "(Optional) A number used to calculate a starting value for the pseudo-random number sequence.")] object seed)
        {
            Faker faker = CreateFaker(locale, seed);

            return faker.Name.JobType();
        }

        [ExcelFunction(Category = "String", Description = "Generates a legit Internet URL avatar from twitter accounts.", HelpTopic = "Generates a legit Internet URL avatar from twitter accounts.")]
        public static object BOGUSAVATAR(
[ExcelArgument("[locale]", Name = "[locale]", Description = "(Optional) The current locale for the dataset.")] object locale,
[ExcelArgument("[seed]", Name = "[seed]", Description = "(Optional) A number used to calculate a starting value for the pseudo-random number sequence.")] object seed)
        {
            Faker faker = CreateFaker(locale, seed);

            return faker.Internet.Avatar();
        }

        [ExcelFunction(Category = "String", Description = "Helper method to pick a random element.", HelpTopic = "Helper method to pick a random element.")]
        public static object BOGUSPICKRANDOM(
[ExcelArgument("values", Name = "values", Description = "A list of elements to pick from.")] object values,
[ExcelArgument("[locale]", Name = "[locale]", Description = "(Optional) The current locale for the dataset.")] object locale,
[ExcelArgument("[seed]", Name = "[seed]", Description = "(Optional) A number used to calculate a starting value for the pseudo-random number sequence.")] object seed)
        {
            Faker faker = CreateFaker(locale, seed);

            object[,] inputArray = (object[,])values;

            List<string> list = FlattenArray(inputArray).Distinct().ToList();

            return faker.PickRandom(list);
        }

        private static object BogusGenerateWithGender(object gender, object locale, object seed, Faker faker, Func<Gender?,string> func)
        {
            try
            {
                string name = func(faker.PickRandom<Gender>());

                string? genderTyped = Optional.Check(gender, null);
                if (genderTyped != null)
                {
                    if (genderTyped.ToLower() == "male")
                    {
                        name = func(Gender.Male);
                    }
                    else if (genderTyped.ToLower() == "female")
                    {
                        name = func(Gender.Female);
                    }
                }

                return name;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        private static Faker CreateFaker(object locale, object seed)
        {
            int typedSeed = Optional.Check(seed, -1);
            if (typedSeed != -1)
            {
                Randomizer.Seed = new Random(typedSeed);
            }

            Faker faker = new Faker("fr");
            string? typedLocatle = Optional.Check(locale, null);
            if (typedLocatle != null)
            {
                faker = new Faker(typedLocatle);
            }

            return faker;
        }

        public static List<string?> FlattenArray(object[,] inputArray)
        {
            List<string?> flattenedList = new List<string?>();

            int rows = inputArray.GetLength(0);
            int columns = inputArray.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    string? value = Optional.Check(inputArray[i, j], null);
                    flattenedList.Add(value);
                }
            }

            return flattenedList;
        }
    }
}
