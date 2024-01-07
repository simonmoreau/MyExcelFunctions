
using ExcelDna.Integration;
using ExcelFunctions.Services;
using Fluid;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Drawing;
using System.IO;

namespace ExcelFunctions
{
    public class Page
    {
        public Page()
        {

        }
        public int Number { get; set; }

        public string Name { get; set; }

    }
        public static class Templating
    {
        [ExcelFunction(Category = "Templating", Description = "Converts the provided value into a JSON string.", HelpTopic = "Converts the provided value into a JSON string.")]
        public static object RENDERTEMPLATE(
            [ExcelArgument("template", Name = "template", Description = "The template.")] string template,
            [ExcelArgument("table", Name = "table", Description = "The table of values to render.")] object values)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type objectType = null;
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out objectType);

                    TemplateOptions.Default.MemberAccessStrategy = new UnsafeMemberAccessStrategy();
                    FluidParser parser = new FluidParser();


                    object model = new { pages = objects };
                    string source = template;

                    if (parser.TryParse(source, out IFluidTemplate? fluidTemplate, out string? error))
                    {
                        TemplateContext context = new TemplateContext(model);

                        string renderedValue = fluidTemplate.Render(context);

                        return renderedValue;
                    }
                    else
                    {
                        return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                    }
                }
                catch (Exception ex)
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }

        [ExcelFunction(Category = "Templating", Description = "Converts the provided value into a JSON string.", HelpTopic = "Converts the provided value into a JSON string.")]
        public static object RENDERTEMPLATETOFILE(
    [ExcelArgument("path", Name = "path", Description = "The file to write to.")] string path,
    [ExcelArgument("template", Name = "template", Description = "The template.")] string template,
    [ExcelArgument("table", Name = "table", Description = "The table of values to render.")] object values)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type objectType = null;
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out objectType);

                    TemplateOptions.Default.MemberAccessStrategy = new UnsafeMemberAccessStrategy();
                    FluidParser parser = new FluidParser();


                    object model = new { pages = objects };
                    string source = template;

                    if (parser.TryParse(source, out IFluidTemplate? fluidTemplate, out string? error))
                    {
                        TemplateContext context = new TemplateContext(model);

                        string renderedValue = fluidTemplate.Render(context);

                        System.IO.File.WriteAllText(path, renderedValue);
                        return $"Text written to {path}";
                    }
                    else
                    {
                        return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                    }
                }
                catch (Exception ex)
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }
    }
}
