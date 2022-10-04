using System;

namespace Rubberduck.Core.WebApi.Model
{
    /// <summary>
    /// Encapsulates the base properties of a data model entity.
    /// </summary>
    public abstract class Entity : IEntity
    {
        /// <summary>
        /// The internal primary key.
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// The timestamp when the entity was first created.
        /// </summary>
        public DateTime DateInserted { get; set; }
        /// <summary>
        /// The timestamp when the entity was last updated.
        /// </summary>
        public DateTime? DateUpdated { get; set; }
    }
}
