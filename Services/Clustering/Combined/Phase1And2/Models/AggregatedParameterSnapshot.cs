using System;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models
{
    /// <summary>
    /// Aggregates per-sleeve parameter snapshots for combined clustering diagnostics and downstream persistence.
    /// </summary>
    public class AggregatedParameterSnapshot
    {
        private readonly Dictionary<int, Dictionary<string, string>> _snapshots = new Dictionary<int, Dictionary<string, string>>();

        public IReadOnlyDictionary<int, Dictionary<string, string>> Snapshots => _snapshots;

        public void Add(int sleeveInstanceId, IDictionary<string, string> parameters)
        {
            if (sleeveInstanceId <= 0 || parameters == null)
            {
                return;
            }

            _snapshots[sleeveInstanceId] = new Dictionary<string, string>(parameters, StringComparer.OrdinalIgnoreCase);
        }

        public Dictionary<string, string> FlattenDistinct()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in _snapshots)
            {
                foreach (var parameter in entry.Value)
                {
                    if (result.ContainsKey(parameter.Key))
                    {
                        continue;
                    }

                    result[parameter.Key] = parameter.Value;
                }
            }

            return result;
        }
    }
}
