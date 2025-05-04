function Tests() {
    Import.UnitTesting.init();

    // ===== Utils =====
    (() =>
        describe('containsKeyword', () => {
            it('end of text', () => {
                assert.equals({
                    expected: true,
                    actual: containsKeyword('test Eng', 'Eng'),
                });
            });

            it('end of text 2', () => {
                assert.equals({
                    expected: true,
                    actual: containsKeyword('Isha Maha Picnic Rus', 'Russian'.slice(0, 3)),
                });
            });
        }))();
}
