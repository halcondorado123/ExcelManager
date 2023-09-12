using OfficeOpenXml;
using System.Reflection.Emit;
using System.Security.Cryptography.Xml;

namespace ExcelManager.Models
{
    public class SubjectsME
    {
            public int?[] Numero { get; set; }
            public string[] Asignatura { get; set; }
            public int?[] Creditos { get; set; }
            public int?[] Semestre { get; set; }
            public double?[] CalificacionNumerica { get; set; }
            public string[] CalificacionLiteral { get; set; }
            public string[] Nivel { get; set; }

        }

        public class GetSubjects
        {
            public static void SubjectsGen()
            {
                // Llamamos al método SubjectGenerator y almacenamos el resultado en una lista
                var subjects = SubjectGenerator();

                // Hacer algo con la lista de asignaturas, como exportar a Excel
            }

            public static List<SubjectsME> SubjectGenerator()
            {
                var subjects = new List<SubjectsME>
            // Crear una lista de asignaturas
            {
            new SubjectsME
            {
                Numero = new int?[]{ 1 },
                Asignatura = new string[] {"LÓGICA Y PENSAMIENTO MATEMÁTICO"},
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 2 },
                Asignatura = new string[] { "PROYECTO DE VIDA" },
                Creditos =  new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 3 },
                Asignatura = new string[] { "INFORMATICA Y CONVERGENCIA TECNOLÓGICA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 4 },
                Asignatura = new string[] { "TERMINOLOGÍA DE LA SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 5 },
                Asignatura = new string[] { "HABILIDADES COMUNICATIVAS" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 6 },
                Asignatura = new string[] { "FUNDAMENTOS DE ADMINISTRACIÓN" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 7 },
                Asignatura = new string[] { "ATENCIÓN AL USUARIO" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 8 },
                Asignatura = new string[] { "CONTABILIDAD BÁSICA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 9 },
                Asignatura = new string[] { "EXPLORAR PARA INVESTIGAR" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 10 },
                Asignatura = new string[] { "MATEMATICA 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 11 },
                Asignatura = new string[] { "CONTABILIDAD DE ENTIDADES DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 12 },
                Asignatura = new string[] { "ADMINISTRACIÓN PÚBLICA DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 13 },
                Asignatura = new string[] { "SISTEMAS DE INFORMACIÓN EN SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 14 },
                Asignatura = new string[] { "FUNDAMENTOS DE ECONOMIA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 15 },
                Asignatura = new string[] { "ELECTIVA DE FORMACION INTEGRAL I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 16 },
                Asignatura = new string[] { "MERCADEO EN EMPRESAS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 17 },
                Asignatura = new string[] { "SALUD Y SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 18 },
                Asignatura = new string[] { "ESTADISTICA I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 19 },
                Asignatura = new string[] { "ANÁLISIS DE LA SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 20 },
                Asignatura = new string[] { "LEGISLACIÓN LABORAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 21 },
                Asignatura = new string[] { "OPCIÓN DE GRADO TÉCNICO" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 22 },
                Asignatura = new string[] { "PRÁCTICA TÉCNICO PROFESIONAL" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 23 },
                Asignatura = new string[] { "ELECTIVA DE PROFUNDIZACIÓN 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 24 },
                Asignatura = new string[] { "GESTIÓN DE LA PRODUCCIÓN DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 25 },
                Asignatura = new string[] { "GESTIÓN DE SUMINISTROS EN SALUD (OPCIÓN PROP)" },
                Creditos = new int?[] { 0 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 26 },
                Asignatura = new string[] { "CÁTEDRA DE PENSAMIENTO CUNISTA I" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },


            new SubjectsME
            {
                Numero = new int?[] { 27 },
                Asignatura = new string[] { "MATERIA EXTRA 1" },
                Creditos = new int?[] { 4 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 28 },
                Asignatura = new string[] { "MATERIA EXTRA 2" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 29 },
                Asignatura = new string[] { "MATERIA EXTRA 3" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

           new SubjectsME
            {
                Numero = new int?[] { 30 },
                Asignatura = new string[] { "MATERIA EXTRA 4" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
                {
                    Numero = new int?[]{ 31 },
                    Asignatura = new string[] {"LÓGICA Y PENSAMIENTO MATEMÁTICO"},
                    Creditos = new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

                new SubjectsME
                {
                    Numero = new int?[] { 32 },
                    Asignatura = new string[] { "PROYECTO DE VIDA" },
                    Creditos =  new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

                new SubjectsME
                {
                    Numero = new int?[] { 33 },
                    Asignatura = new string[] { "INFORMATICA Y CONVERGENCIA TECNOLÓGICA" },
                    Creditos = new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

            new SubjectsME
            {
                Numero = new int?[] { 34 },
                Asignatura = new string[] { "TERMINOLOGÍA DE LA SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 35 },
                Asignatura = new string[] { "HABILIDADES COMUNICATIVAS" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 36 },
                Asignatura = new string[] { "FUNDAMENTOS DE ADMINISTRACIÓN" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 37 },
                Asignatura = new string[] { "ATENCIÓN AL USUARIO" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 38 },
                Asignatura = new string[] { "CONTABILIDAD BÁSICA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 39 },
                Asignatura = new string[] { "EXPLORAR PARA INVESTIGAR" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 40 },
                Asignatura = new string[] { "MATEMATICA 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 41 },
                Asignatura = new string[] { "CONTABILIDAD DE ENTIDADES DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 42 },
                Asignatura = new string[] { "ADMINISTRACIÓN PÚBLICA DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 43 },
                Asignatura = new string[] { "SISTEMAS DE INFORMACIÓN EN SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 44 },
                Asignatura = new string[] { "FUNDAMENTOS DE ECONOMIA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 45 },
                Asignatura = new string[] { "ELECTIVA DE FORMACION INTEGRAL I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 46 },
                Asignatura = new string[] { "MERCADEO EN EMPRESAS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 47 },
                Asignatura = new string[] { "SALUD Y SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 48 },
                Asignatura = new string[] { "ESTADISTICA I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 49 },
                Asignatura = new string[] { "ANÁLISIS DE LA SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 50 },
                Asignatura = new string[] { "LEGISLACIÓN LABORAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 51 },
                Asignatura = new string[] { "OPCIÓN DE GRADO TÉCNICO" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 52 },
                Asignatura = new string[] { "PRÁCTICA TÉCNICO PROFESIONAL" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 53 },
                Asignatura = new string[] { "ELECTIVA DE PROFUNDIZACIÓN 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 54 },
                Asignatura = new string[] { "GESTIÓN DE LA PRODUCCIÓN DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 55 },
                Asignatura = new string[] { "GESTIÓN DE SUMINISTROS EN SALUD (OPCIÓN PROP)" },
                Creditos = new int?[] { 0 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 56 },
                Asignatura = new string[] { "CÁTEDRA DE PENSAMIENTO CUNISTA I" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },


            new SubjectsME
            {
                Numero = new int?[] { 57 },
                Asignatura = new string[] { "MATERIA EXTRA 1" },
                Creditos = new int?[] { 4 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 58 },
                Asignatura = new string[] { "MATERIA EXTRA 2" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 59 },
                Asignatura = new string[] { "MATERIA EXTRA 3" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

           new SubjectsME
            {
                Numero = new int?[] { 60 },
                Asignatura = new string[] { "MATERIA EXTRA 4" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
                {
                    Numero = new int?[]{ 61 },
                    Asignatura = new string[] {"LÓGICA Y PENSAMIENTO MATEMÁTICO"},
                    Creditos = new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

                new SubjectsME
                {
                    Numero = new int?[] { 62 },
                    Asignatura = new string[] { "PROYECTO DE VIDA" },
                    Creditos =  new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

                new SubjectsME
                {
                    Numero = new int?[] { 63 },
                    Asignatura = new string[] { "INFORMATICA Y CONVERGENCIA TECNOLÓGICA" },
                    Creditos = new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

            new SubjectsME
            {
                Numero = new int?[] { 64 },
                Asignatura = new string[] { "TERMINOLOGÍA DE LA SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 65 },
                Asignatura = new string[] { "HABILIDADES COMUNICATIVAS" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 66 },
                Asignatura = new string[] { "FUNDAMENTOS DE ADMINISTRACIÓN" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 67 },
                Asignatura = new string[] { "ATENCIÓN AL USUARIO" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 68 },
                Asignatura = new string[] { "CONTABILIDAD BÁSICA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 69 },
                Asignatura = new string[] { "EXPLORAR PARA INVESTIGAR" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 70 },
                Asignatura = new string[] { "MATEMATICA 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 71 },
                Asignatura = new string[] { "CONTABILIDAD DE ENTIDADES DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 72 },
                Asignatura = new string[] { "ADMINISTRACIÓN PÚBLICA DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 73 },
                Asignatura = new string[] { "SISTEMAS DE INFORMACIÓN EN SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 74 },
                Asignatura = new string[] { "FUNDAMENTOS DE ECONOMIA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 75 },
                Asignatura = new string[] { "ELECTIVA DE FORMACION INTEGRAL I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 76 },
                Asignatura = new string[] { "MERCADEO EN EMPRESAS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 77 },
                Asignatura = new string[] { "SALUD Y SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 78 },
                Asignatura = new string[] { "ESTADISTICA I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 79 },
                Asignatura = new string[] { "ANÁLISIS DE LA SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 80 },
                Asignatura = new string[] { "LEGISLACIÓN LABORAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 81 },
                Asignatura = new string[] { "OPCIÓN DE GRADO TÉCNICO" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 82 },
                Asignatura = new string[] { "PRÁCTICA TÉCNICO PROFESIONAL" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 83 },
                Asignatura = new string[] { "ELECTIVA DE PROFUNDIZACIÓN 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 84 },
                Asignatura = new string[] { "GESTIÓN DE LA PRODUCCIÓN DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 85 },
                Asignatura = new string[] { "GESTIÓN DE SUMINISTROS EN SALUD (OPCIÓN PROP)" },
                Creditos = new int?[] { 0 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 86 },
                Asignatura = new string[] { "CÁTEDRA DE PENSAMIENTO CUNISTA I" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },


            new SubjectsME
            {
                Numero = new int?[] { 87 },
                Asignatura = new string[] { "MATERIA EXTRA 1" },
                Creditos = new int?[] { 4 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 88 },
                Asignatura = new string[] { "MATERIA EXTRA 2" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 89 },
                Asignatura = new string[] { "MATERIA EXTRA 3" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

           new SubjectsME
            {
                Numero = new int?[] { 90 },
                Asignatura = new string[] { "MATERIA EXTRA 4" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 91 },
                Asignatura = new string[] { "QUIMICA AEROESPACIAL tRIGONOMETRICA IV" },
                Creditos = new int?[] { 8 },
                Semestre = new int?[] { 7 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

                        new SubjectsME
            {
                Numero = new int?[] { 92 },
                Asignatura = new string[] { "PRÁCTICA TÉCNICO PROFESIONAL" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 93 },
                Asignatura = new string[] { "ELECTIVA DE PROFUNDIZACIÓN 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 94 },
                Asignatura = new string[] { "GESTIÓN DE LA PRODUCCIÓN DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 95 },
                Asignatura = new string[] { "GESTIÓN DE SUMINISTROS EN SALUD (OPCIÓN PROP)" },
                Creditos = new int?[] { 0 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 96 },
                Asignatura = new string[] { "CÁTEDRA DE PENSAMIENTO CUNISTA I" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },


            new SubjectsME
            {
                Numero = new int?[] { 97 },
                Asignatura = new string[] { "MATERIA EXTRA 1" },
                Creditos = new int?[] { 4 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 98 },
                Asignatura = new string[] { "MATERIA EXTRA 2" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 99 },
                Asignatura = new string[] { "MATERIA EXTRA 3" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

           new SubjectsME
            {
                Numero = new int?[] { 100 },
                Asignatura = new string[] { "MATERIA EXTRA 4" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 4 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new SubjectsME
            {
                Numero = new int?[] { 101 },
                Asignatura = new string[] { "ESTUDIOS CALCULOALGEBRAICO MULTIDIMENSIONAL" },
                Creditos = new int?[] { 8 },
                Semestre = new int?[] { 7 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            }

        };


                return subjects;
            }
        }

}

