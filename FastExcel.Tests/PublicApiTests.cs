using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Guards the shape of the public API.
    ///
    /// Every other test in this project answers "does the library behave correctly?". This one
    /// answers a different question: "will a program compiled against the previous release still
    /// start?". Those are not the same. Adding a parameter to a public constructor keeps every
    /// behavioural test green while raising <see cref="MissingMethodException"/> in every
    /// already-shipped consumer, because the signature is baked into their compiled call site.
    ///
    /// That is not hypothetical here. The accepted fix for #75 (read-only streams) adds a
    /// parameter to a public constructor, and the package has ~780,000 downloads across ~124
    /// dependent repositories.
    ///
    /// This test renders the public surface as sorted text and compares it to a committed
    /// baseline. Any addition, removal or signature change fails it. A failure is not by itself
    /// a bug — it is a prompt to decide, deliberately, whether the change is additive (safe) or
    /// breaking (needs a major version), and to update the baseline in the same commit so the
    /// diff records that decision.
    /// </summary>
    public class PublicApiTests
    {
        private const string BaselineFileName = "PublicApi.approved.txt";

        [Fact]
        public void ThePublicApiSurfaceHasNotChanged()
        {
            var current = RenderPublicApi(typeof(FastExcel).Assembly);
            var baselinePath = LocateBaseline();

            if (!File.Exists(baselinePath))
            {
                File.WriteAllText(baselinePath, current);
                Assert.Fail(
                    "No API baseline existed, so one was written to " + baselinePath +
                    ". Review it and commit it - from now on it is the contract.");
            }

            var approved = Normalise(File.ReadAllText(baselinePath));
            if (approved == Normalise(current))
            {
                return;
            }

            var receivedPath = Path.Combine(Path.GetDirectoryName(baselinePath), "PublicApi.received.txt");
            File.WriteAllText(receivedPath, current);

            Assert.Fail(
                "The public API surface changed.\n\n" +
                Describe(approved, Normalise(current)) +
                "\nFull rendering written to " + receivedPath + "\n\n" +
                "If the change is deliberate: copy it over " + BaselineFileName + " and commit both together.\n" +
                "Removed or altered members break every already-compiled consumer and need a major version.");
        }

        /// <summary>
        /// Every exported type should be obtainable by a consumer — through a public constructor,
        /// or as the return or property type of some public member. A type that is neither is
        /// public by accident: it costs binary-compatibility obligations and gives nothing back.
        ///
        /// <c>SharedStrings</c> is the live example, and it is the property #94 points at.
        /// It is <c>public</c> and exposes a settable <c>ReadWriteMode</c> that changes how every
        /// later read resolves strings, but its constructor is <c>internal</c> and no public
        /// member hands one out — so the switch is documented, reachable in the API listing, and
        /// impossible to actually touch. The write path toggles it around the row loop without a
        /// try/finally, so an exception mid-write leaves it stuck on for the object's lifetime.
        ///
        /// Either it should be internal, or it should be reachable and tested. Right now it is
        /// the worst of both.
        /// </summary>
        [Fact]
        public void EveryExportedTypeIsReachableThroughThePublicApi()
        {
            var assembly = typeof(FastExcel).Assembly;
            var exported = assembly.GetExportedTypes();

            // Everything a consumer could ever be handed by a public member.
            var obtainable = new HashSet<Type>();
            foreach (var type in exported)
            {
                foreach (var member in type.GetMembers(BindingFlags.Public | BindingFlags.Instance |
                                                       BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    if (member is MethodInfo method) obtainable.Add(Unwrap(method.ReturnType));
                    else if (member is PropertyInfo property) obtainable.Add(Unwrap(property.PropertyType));
                    else if (member is FieldInfo field) obtainable.Add(Unwrap(field.FieldType));
                }
            }

            var unreachable = exported
                .Where(t => !t.IsAbstract || !t.IsSealed)                 // static classes are reachable by name
                .Where(t => t.GetConstructors().Length == 0)              // no public constructor
                .Where(t => !obtainable.Contains(t))                      // and never handed out
                .Where(t => !typeof(Exception).IsAssignableFrom(t))       // exceptions are caught, not constructed
                .Where(t => !typeof(Attribute).IsAssignableFrom(t))       // attributes are applied, not constructed
                .Select(t => t.Name)
                .OrderBy(n => n, StringComparer.Ordinal)
                .ToList();

            KnownBug.StillBroken("#94",
                "every exported type can be obtained through the public API; today " +
                "SharedStrings is public with a settable ReadWriteMode that alters every later " +
                "read, yet no public member returns one and its constructor is internal",
                () => Assert.True(unreachable.Count == 0,
                    "exported but unobtainable: " + string.Join(", ", unreachable)));
        }

        /// <summary>
        /// The add-and-delete-worksheet feature is not implemented, and the machinery written for
        /// it is unreachable.
        ///
        /// <c>FastExcel</c> carries two private fields, <c>AddWorksheets</c> and
        /// <c>DeleteWorksheets</c>, that are read in nine places and assigned in none. Every
        /// branch guarded by them is dead: roughly 180 lines across <c>UpdateRelations</c>,
        /// <c>UpdateWorkbook</c>, <c>RenameAndRebildWorksheetProperties</c> and
        /// <c>UpdateContentTypes</c>, plus <c>Worksheet.ValidateNewWorksheet</c>,
        /// <c>Worksheet.AddSettings</c> and the whole <c>WorksheetAddSettings</c> type.
        ///
        /// This is why overall coverage stops in the low eighties rather than approaching 100%:
        /// no test can reach that code, because nothing can. Worth knowing, rather than filing
        /// under "we should test more".
        ///
        /// The test fails if a public way to add or remove a sheet ever appears - at which point
        /// the machinery becomes reachable, needs real tests, and this note should go.
        /// </summary>
        [Fact]
        public void ThereIsNoPublicWayToAddOrRemoveAWorksheet()
        {
            var members = typeof(FastExcel).Assembly.GetExportedTypes()
                .SelectMany(t => t.GetMembers(BindingFlags.Public | BindingFlags.Instance |
                                              BindingFlags.Static | BindingFlags.DeclaredOnly))
                .Select(m => m.Name)
                .ToList();

            foreach (var verb in new[] { "AddWorksheet", "DeleteWorksheet", "RemoveWorksheet", "InsertWorksheet" })
            {
                Assert.DoesNotContain(members,
                    n => n.IndexOf(verb, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            // And the switches that would turn the machinery on are private, with no path to set them.
            var privateProperties = typeof(FastExcel)
                .GetProperties(BindingFlags.NonPublic | BindingFlags.Instance)
                .Select(p => p.Name)
                .ToList();

            Assert.Contains("AddWorksheets", privateProperties);
            Assert.Contains("DeleteWorksheets", privateProperties);
        }

        /// <summary>Looks through arrays, nullables and one level of generics to the type inside.</summary>
        private static Type Unwrap(Type type)
        {
            if (type.IsByRef || type.IsArray) return Unwrap(type.GetElementType());
            var nullable = Nullable.GetUnderlyingType(type);
            if (nullable != null) return Unwrap(nullable);
            return type;
        }

        /// <summary>Reports what was added and removed, so the failure reads without a diff tool.</summary>
        private static string Describe(string approved, string current)
        {
            var before = new HashSet<string>(approved.Split('\n').Where(l => l.Length > 0));
            var after = new HashSet<string>(current.Split('\n').Where(l => l.Length > 0));

            var removed = before.Except(after).OrderBy(l => l, StringComparer.Ordinal).ToList();
            var added = after.Except(before).OrderBy(l => l, StringComparer.Ordinal).ToList();

            var report = new StringBuilder();
            if (removed.Count > 0)
            {
                report.AppendLine("REMOVED OR CHANGED (" + removed.Count + ") - breaks compiled consumers:");
                foreach (var line in removed) report.AppendLine("  - " + line);
            }
            if (added.Count > 0)
            {
                report.AppendLine("ADDED (" + added.Count + ") - additive, safe for existing consumers:");
                foreach (var line in added) report.AppendLine("  + " + line);
            }
            return report.ToString();
        }

        private static string Normalise(string text) => text.Replace("\r\n", "\n").Trim();

        /// <summary>
        /// Finds the baseline in the source tree rather than the output folder, so a failing run
        /// leaves the received file next to the file a developer actually has to edit.
        /// </summary>
        private static string LocateBaseline()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "FastExcel.Tests.csproj")))
            {
                directory = directory.Parent;
            }

            Assert.True(directory != null,
                "could not locate the test project directory from " + AppContext.BaseDirectory);

            return Path.Combine(directory.FullName, BaselineFileName);
        }

        // ------------------------------------------------------------------ rendering

        /// <summary>
        /// Renders every externally reachable member as one sorted line. Sorting is ordinal so the
        /// output does not depend on the machine's culture, and members come from metadata rather
        /// than source order, so reordering a source file does not fail the test.
        /// </summary>
        private static string RenderPublicApi(Assembly assembly)
        {
            var lines = new List<string>();

            foreach (var type in assembly.GetExportedTypes())
            {
                lines.Add(DescribeType(type));

                const BindingFlags flags = BindingFlags.Public | BindingFlags.NonPublic |
                                           BindingFlags.Instance | BindingFlags.Static |
                                           BindingFlags.DeclaredOnly;

                foreach (var member in type.GetMembers(flags))
                {
                    // protected members are part of the surface for anyone deriving from the type.
                    if (!IsExternallyVisible(member)) continue;

                    var rendered = DescribeMember(type, member);
                    if (rendered != null) lines.Add(rendered);
                }
            }

            lines.Sort(StringComparer.Ordinal);
            return string.Join(Environment.NewLine, lines) + Environment.NewLine;
        }

        private static bool IsExternallyVisible(MemberInfo member)
        {
            if (member is MethodBase method)
            {
                return method.IsPublic || method.IsFamily || method.IsFamilyOrAssembly;
            }
            if (member is FieldInfo field)
            {
                return field.IsPublic || field.IsFamily || field.IsFamilyOrAssembly;
            }
            if (member is PropertyInfo property)
            {
                return new[] { property.GetMethod, property.SetMethod }
                    .Where(a => a != null)
                    .Any(a => a.IsPublic || a.IsFamily || a.IsFamilyOrAssembly);
            }
            if (member is EventInfo evt)
            {
                return evt.AddMethod != null && (evt.AddMethod.IsPublic || evt.AddMethod.IsFamily);
            }
            if (member is Type nested)
            {
                return nested.IsNestedPublic || nested.IsNestedFamily;
            }
            return false;
        }

        private static string DescribeType(Type type)
        {
            var kind = type.IsEnum ? "enum"
                     : type.IsInterface ? "interface"
                     : type.IsValueType ? "struct"
                     : "class";

            var modifiers = new List<string>();
            if (type.IsAbstract && type.IsSealed)
            {
                modifiers.Add("static");
            }
            else
            {
                if (type.IsAbstract && !type.IsInterface) modifiers.Add("abstract");
                if (type.IsSealed && !type.IsEnum && !type.IsValueType) modifiers.Add("sealed");
            }

            var bases = new List<string>();
            if (type.BaseType != null && type.BaseType != typeof(object) &&
                type.BaseType != typeof(ValueType) && type.BaseType != typeof(Enum))
            {
                bases.Add(Name(type.BaseType));
            }
            bases.AddRange(type.GetInterfaces().Select(Name).OrderBy(n => n, StringComparer.Ordinal));

            var declaration = "TYPE " + string.Join(" ", modifiers.Concat(new[] { kind })) + " " + Name(type);
            return bases.Count > 0 ? declaration + " : " + string.Join(", ", bases) : declaration;
        }

        private static string DescribeMember(Type type, MemberInfo member)
        {
            if (member is ConstructorInfo ctor)
            {
                return "  " + Name(type) + "..ctor(" + Parameters(ctor) + ")";
            }

            if (member is MethodInfo method)
            {
                if (IsAccessor(method)) return null;
                return "  " + Modifiers(method) + Name(type) + "." + method.Name +
                       "(" + Parameters(method) + ") : " + Name(method.ReturnType);
            }

            if (member is PropertyInfo property)
            {
                var accessors = new List<string>();
                if (property.GetMethod != null &&
                    (property.GetMethod.IsPublic || property.GetMethod.IsFamily)) accessors.Add("get");
                if (property.SetMethod != null &&
                    (property.SetMethod.IsPublic || property.SetMethod.IsFamily)) accessors.Add("set");

                var index = property.GetIndexParameters();
                var name = index.Length > 0
                    ? "this[" + string.Join(", ", index.Select(p => Name(p.ParameterType))) + "]"
                    : property.Name;

                return "  " + Name(type) + "." + name + " : " + Name(property.PropertyType) +
                       " { " + string.Join("; ", accessors) + "; }";
            }

            if (member is FieldInfo field)
            {
                var constness = field.IsLiteral ? "const " : field.IsStatic ? "static " : string.Empty;
                return "  " + constness + Name(type) + "." + field.Name + " : " + Name(field.FieldType);
            }

            if (member is EventInfo evt)
            {
                return "  event " + Name(type) + "." + evt.Name + " : " + Name(evt.EventHandlerType);
            }

            return null;
        }

        /// <summary>Property and event accessors are reported through their owner, not twice.</summary>
        private static bool IsAccessor(MethodInfo method)
        {
            return method.IsSpecialName &&
                   (method.Name.StartsWith("get_", StringComparison.Ordinal) ||
                    method.Name.StartsWith("set_", StringComparison.Ordinal) ||
                    method.Name.StartsWith("add_", StringComparison.Ordinal) ||
                    method.Name.StartsWith("remove_", StringComparison.Ordinal));
        }

        private static string Modifiers(MethodInfo method)
        {
            var parts = new List<string>();
            if (method.IsStatic) parts.Add("static");
            if (method.IsAbstract) parts.Add("abstract");
            else if (method.IsVirtual && !method.IsFinal) parts.Add("virtual");
            return parts.Count > 0 ? string.Join(" ", parts) + " " : string.Empty;
        }

        private static string Parameters(MethodBase method)
        {
            return string.Join(", ", method.GetParameters().Select(p =>
            {
                var prefix = p.IsOut ? "out " : p.ParameterType.IsByRef ? "ref " : string.Empty;
                var suffix = p.IsOptional ? " = default" : string.Empty;
                return prefix + Name(p.ParameterType) + suffix;
            }));
        }

        /// <summary>Short, stable type names - full names would make the baseline unreadable.</summary>
        private static string Name(Type type)
        {
            if (type.IsByRef) return Name(type.GetElementType());
            if (type.IsArray) return Name(type.GetElementType()) + "[]";

            var nullable = Nullable.GetUnderlyingType(type);
            if (nullable != null) return Name(nullable) + "?";

            if (type.IsGenericType)
            {
                var stem = type.Name.Substring(0, type.Name.IndexOf('`'));
                var args = string.Join(", ", type.GetGenericArguments().Select(Name));
                return stem + "<" + args + ">";
            }

            switch (type.FullName)
            {
                case "System.Void": return "void";
                case "System.Boolean": return "bool";
                case "System.Int32": return "int";
                case "System.Int64": return "long";
                case "System.Double": return "double";
                case "System.String": return "string";
                case "System.Object": return "object";
                default: return type.Name;
            }
        }
    }
}
