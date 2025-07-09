namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Encapsulates clearance values for each side of a rectangular sleeve.
    /// </summary>
    public class SleeveClearance
    {
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }

        public SleeveClearance(double left, double right, double top, double bottom)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
        }

        /// <summary>
        /// Standard clearance: 50mm on all sides.
        /// </summary>
        public static SleeveClearance Standard() => new SleeveClearance(50, 50, 50, 50);

        /// <summary>
        /// Insulated clearance: 25mm on all sides.
        /// </summary>
        public static SleeveClearance Insulated() => new SleeveClearance(25, 25, 25, 25);

        /// <summary>
        /// Damper clearance: 100mm on connector side, 50mm elsewhere.
        /// </summary>
        /// <param name="connectorSide">"Top", "Right", "Bottom", or "Left"</param>
        public static SleeveClearance Damper(string connectorSide)
        {
            switch (connectorSide)
            {
                case "Top": return new SleeveClearance(50, 50, 100, 50);
                case "Right": return new SleeveClearance(50, 100, 50, 50);
                case "Bottom": return new SleeveClearance(50, 50, 50, 100);
                case "Left": return new SleeveClearance(100, 50, 50, 50);
                default: return new SleeveClearance(50, 100, 50, 50); // Default to right
            }
        }
    }
}
