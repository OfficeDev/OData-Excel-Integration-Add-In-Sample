var UT;
(function (UT) {
    describe("Animation", function () {
        describe("TransformingCalculator_3d", function () {
            var precision = 10;
            it("returns the right value", function () {
                var height = 2;
                var calculator = new Animation.TransformingCalculator_3d(height);
                var angle = calculator.getSunkAngle(1);
                expect(angle).toBeCloseTo(0, precision);
                angle = calculator.getSunkAngle(0);
                expect(angle).toBeCloseTo(Math.PI / 2, precision);
            });
        });
    });
})(UT || (UT = {}));
//# sourceMappingURL=ut_Animation.js.map