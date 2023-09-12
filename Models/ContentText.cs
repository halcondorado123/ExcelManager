namespace ExcelManager.Models
{
    public class ContentText
    {

        public string TitleCondition() { return "Aclaraciones"; }
        public string TitleImportantCondition() { return "Manifestación expresa del estudiante"; }
        public string TextInConstancy() { return "En Constancia de lo anterior firman:"; }

        public string[] TotalCreditsValue()
        {
            return new string[]
            {
                "APRO",
                "PEN",
                "TOTAL CRED"
            };
        }


        public string[] HeaderTableCells()
        {
            return new string[]
            {
            "No",
            "ASIGNATURA Y/O CRÉDITO HOMOLOGADO",
            "SISTEMA",
            "CALIFICACIÓN NUMERICA",
            "CALIFICACION LITERAL",
            "NIVEL",
            };
        }

        public string[] HeaderTableSubCells()
        {
            return new string[]
            {
                "Créditos",
                "Semestre",
            };
        }

        public string[] GetLabelsTotalCredits()
        {
            return new string[]
            {
            "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TÉCNICO PROFESIONAL",
            "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TECNOLÓGICO",
            "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL PROFESIONAL"
            };
        }


        public List<string> ConditionLabels()
        {
            return new List<string>
            {
            "Los créditos académicos faltantes para cumplir a cabalidad con la oferta académica del nivel técnico profesional y/o tecnológico y/o " +
            "profesional deben ser cursados y aprobados conforme las reglamentaciones institucionales vigentes.",
            "La Escuela de Ciencias Administrativas de la Corporación Unificada Nacional -CUN, reconoce las asignaturas del programa de  Administración " +
            "de Servicios de Salud   del nivel técnico ( 1 - 3  semestre) para que dar continuidad a su proceso de formación académica a partir de las" +
            "asignaturas de nivel tecnológico correspondientes a (  4 - 5  semestre) Y profesional correspondientes a ( 6 - 9  semestre)",
            "La prueba TyT del ciclo técnico es homologable en la institución. El estudiante deberá presentar la prueba saber TyT para el ciclo " +
             "tecnológico y Saber PRO para el ciclo profesional.",
            "Teniendo en cuenta que el plan de estudios vigente del programa  Administración de Servicios de Salud  no incluye los respectivos niveles " +
            "de inglés requeridos para obtener las diferentes titulaciones, el estudiante deberá garantizar lo pertinente al momento de radicar su solicitud de " +
            "grado, para ello se cuenta con la oferta del centro de Idiomas de la institución."
            };
        }

        public List<string> ImportantConditionLabels()
        {
            return new List<string>
            {
            "Con el presente documento manifiesto expresamente y sin que medie ninguna clase de vicio o limitación a mi consentimiento, mi plena " +
            "conformidad con las asignaturas y/o créditos reconocidos u homologados para mi ingreso al nivel técnico profesional y/o tecnológico",
            " y/o profesional del programa  Administración de Servicios de Salud  de la Corporación Unificada Nacional de Educación Superior CUN." +
            " Las competencias que considere me hagan falta, del ciclo técnico, las podré realizar voluntariamente a través de tutorías en cada área" +
            " transversal o del programa, talleres nivelatorios y/o participando como asistente a clases sin que estos generen nota alguna y solicitando" +
            " previamente el ingreso ",
            "a la clase o tutoría."
            };
        }

        public string[] AddProgramManagerInfo()
        {
            return new string[]
            {
            "Líder de Programa: ",
            "Nombre: "
            };
        }

        public string[] AddProgramStudentInfo()
        {
            return new string[]
            {
            "Estudiante:	",
            "Nombre: ",
            "Doc de Identidad: "
            };
        }


    }
}
