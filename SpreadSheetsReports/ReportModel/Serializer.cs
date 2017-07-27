namespace SpreadSheetsReports.ReportModel
{
    using Evaluator;
    using System;
    using System.CodeDom;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.Serialization;

    public class Serializer
    {
        public static ReportDefinition Deserialize(string path)
        {
            DataContractSerializer reader = GetSerializer();

            using (var file = File.OpenRead(path))
            {
                return (ReportDefinition)reader.ReadObject(file);
            }
        }

        public static void Serialize(ReportDefinition content, string path)
        {
            DataContractSerializer writer = GetSerializer();

            using (var file = File.Create(path))
            {
                writer.WriteObject(file, content);
            }
        }

        private static DataContractSerializer GetSerializer()
        {
            List<Assembly> allAssemblies = new List<Assembly>();
            string binPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            foreach (string dll in Directory.GetFiles(binPath, "*.dll"))
            {
                allAssemblies.Add(Assembly.LoadFile(dll));
            }

            var type = typeof(ISheetGenerator);
            var types = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && !p.IsInterface).ToList().AsEnumerable();
            type = typeof(PropertyBindingBase);
            types = types.Concat(AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && !p.IsInterface).ToList().AsEnumerable());
            type = typeof(IRowCollectionGenerator);
            types = types.Concat(AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && !p.IsInterface).ToList().AsEnumerable());
            type = typeof(Row);
            types = types.Concat(AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && !p.IsInterface).ToList().AsEnumerable());
            type = typeof(IEvaluator);
            types = types.Concat(AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && !p.IsInterface).ToList().AsEnumerable());
            type = typeof(Cell);
            types = types.Concat(AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && !p.IsInterface).ToList().AsEnumerable());
            DataContractSerializer serializer = new DataContractSerializer(typeof(ReportDefinition), types, int.MaxValue, false, true, new Surrogate());
            return serializer;
        }

        private class Surrogate : IDataContractSurrogate
        {
            public object GetCustomDataToExport(Type clrType, Type dataContractType)
            {
                return null;
            }

            public object GetCustomDataToExport(MemberInfo memberInfo, Type dataContractType)
            {
                return null;
            }

            public Type GetDataContractType(Type type)
            {
                return type;
            }

            public object GetDeserializedObject(object obj, Type targetType)
            {
                return obj;
            }

            public void GetKnownCustomDataTypes(Collection<Type> customDataTypes)
            {
            }

            public object GetObjectToSerialize(object obj, Type targetType)
            {
                return obj;
            }

            public Type GetReferencedTypeOnImport(string typeName, string typeNamespace, object customData)
            {
                return null;
            }

            public CodeTypeDeclaration ProcessImportedType(CodeTypeDeclaration typeDeclaration, CodeCompileUnit compileUnit)
            {
                return typeDeclaration;
            }
        }
    }
}
