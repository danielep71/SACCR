"""Regression cases for post-merge reviews on PRs 21 and 24."""
import unittest

import check_vba_conditionals as conditionals
import check_vba_jumps as jumps
import check_vba_public_api as public_api


class ReviewRegressionTests(unittest.TestCase):
    def test_exclusive_labels_are_not_duplicates(self) -> None:
        source = ('Public Sub Probe()\n#If Win64 Then\nGoTo Done\nDone:\n'
                  '#Else\nGoTo Done\nDone:\n#End If\nEnd Sub\n')
        self.assertEqual(jumps.analyze_component('Probe.bas', source), [])

    def test_label_in_opposite_branch_cannot_satisfy_jump(self) -> None:
        source = ('Public Sub Probe()\n#If Win64 Then\nGoTo Done\n'
                  '#Else\nDone:\n#End If\nEnd Sub\n')
        findings = jumps.analyze_component('Probe.bas', source)
        self.assertEqual(len(findings), 1)
        self.assertIn('Done', findings[0]['message'])
        self.assertEqual(findings[0]['line'], 3)
        self.assertEqual(len(findings[0]['environments']), 1)

    def test_unknown_jump_condition_fails_closed(self) -> None:
        findings = jumps.analyze_component('Probe.bas', '#If Unknown Then\n#End If\n')
        self.assertTrue(findings)

    def test_same_active_branch_still_rejects_duplicate_labels(self) -> None:
        source = 'Sub Probe()\n#If VBA7 Then\nDone:\nDone:\n#End If\nEnd Sub\n'
        self.assertTrue(jumps.analyze_component('Probe.bas', source))

    def test_ptrsafe_in_string_or_name_is_not_modifier(self) -> None:
        for declaration in (
            'Private Declare Function F Lib "PtrSafe" () As Long',
            'Private Declare Function PtrSafe Lib "kernel32" () As Long',
            'Private Declare Function F Lib "kernel32" Alias "PtrSafe" () As Long',
        ):
            with self.subTest(declaration=declaration):
                self.assertTrue(conditionals.analyze_component('Probe.bas', declaration))

    def test_real_ptrsafe_modifier_and_legacy_branch_pass(self) -> None:
        source = ('#If VBA7 Then\nPrivate Declare PtrSafe Function F Lib "x" () As Long\n'
                  '#Else\nPrivate Declare Function F Lib "x" () As Long\n#End If\n')
        self.assertEqual(conditionals.analyze_component('Probe.bas', source), [])

    def test_colon_hidden_public_declaration_is_rejected(self) -> None:
        source = ('Attribute VB_Name = "Facade"\nOption Explicit\n'
                  'Public Visible As Long: Public Hidden As Long\n')
        manifest = ['Facade\tVariable\tVisible',
                    '# SIG\tFacade\tVariable\tVisible\tPublic Visible As Long']
        result = public_api.fixture(source, manifest)
        self.assertEqual(result['status'], 'fail')
        self.assertTrue(any('Hidden' in str(f) for f in result['findings']))
        manifest += ['Facade\tVariable\tHidden',
                     '# SIG\tFacade\tVariable\tHidden\tPublic Hidden As Long']
        self.assertEqual(public_api.fixture(source, manifest)['status'], 'pass')

    def test_statement_split_preserves_strings_and_named_arguments(self) -> None:
        code = 'Public Const Text As String = "a:""b:c""": Call F(value:=1)'
        self.assertEqual(public_api.split_statements(code),
                         ['Public Const Text As String = "a:""b:c"""', 'Call F(value:=1)'])
        self.assertEqual(public_api.split_statements('Public A As Long: Rem ignored: Public B As Long'),
                         ['Public A As Long'])

    def test_statement_split_preserves_date_literals(self) -> None:
        self.assertEqual(public_api.split_statements('Public Const T As Date = #12:30:00#: Public X As Long'),
                         ['Public Const T As Date = #12:30:00#', 'Public X As Long'])

    def test_colon_procedures_and_continuations(self) -> None:
        source = ('Attribute VB_Name = "Facade"\nOption Explicit\n'
                  'Public Sub A(): End Sub: Public Sub B(): End Sub\n'
                  'Public X As Long: _\nPublic Y As Long\n')
        declarations, errors = public_api.parse_component('src/modules/Facade.bas', source, True)
        self.assertFalse(errors)
        self.assertEqual({item['name'] for item in declarations}, {'A', 'B', 'X', 'Y'})


if __name__ == '__main__':
    unittest.main()
