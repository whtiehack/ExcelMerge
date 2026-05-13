using System;
using FastWpfGrid;
using ExcelMerge.GUI.Models;

namespace ExcelMerge.GUI.Services
{
    /// <summary>
    /// Encapsulates diff-navigation logic: given a grid and a strategy for finding
    /// the target cell, moves the grid's CurrentCell to that target.
    /// </summary>
    public static class DiffNavigator
    {
        /// <summary>
        /// Navigates a grid's CurrentCell using a model-level lookup function.
        /// </summary>
        /// <param name="grid">The grid whose CurrentCell should be updated.</param>
        /// <param name="getTarget">
        /// A function that receives the grid's DiffGridModel and the current cell address,
        /// and returns the target cell address (or Empty if none found).
        /// </param>
        /// <returns>True if navigation succeeded (target found), false otherwise.</returns>
        public static bool Navigate(
            FastGridControl grid,
            Func<DiffGridModel, FastGridCellAddress, FastGridCellAddress> getTarget)
        {
            var model = grid?.Model as DiffGridModel;
            if (model == null) return false;

            var current = grid.CurrentCell.IsEmpty
                ? FastGridCellAddress.Zero
                : grid.CurrentCell;

            var target = getTarget(model, current);
            if (!target.IsEmpty)
            {
                grid.CurrentCell = target;
                return true;
            }
            return false;
        }
    }
}
