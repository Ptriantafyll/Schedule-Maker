# Flutter & Dart Testing Guide

This document is a practical reference for testing Flutter and Dart applications in Schedule-Maker. It covers core concepts, patterns, conventions, and examples that we use across domain models, data sources, repositories, ViewModels, and UI widgets.

---

## 1. Test File Setup & Conventions

- **File Location**: All tests reside under the `test/` directory, mirroring the structure of `lib/`.
  - Production code: `lib/features/auth/domain/models/user_role.dart`
  - Test code: `test/features/auth/domain/models/user_role_test.dart`
- **File Naming**: Test files **must end with `_test.dart`**. The `flutter test` command will ignore any file not matching this suffix.
- **Running Tests**:
  - Run all tests: `flutter test`
  - Run a specific file: `flutter test test/features/auth/domain/models/user_role_test.dart`
  - Run tests with name filter: `flutter test --name "fromString"`

---

## 2. Anatomy of a Dart Test

Every test file is an executable Dart program:

```dart
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

void main() {
  group('UserRole', () {
    test('parses super_admin correctly', () {
      // 1. Arrange & Act
      final role = UserRole.fromString('super_admin');

      // 2. Assert
      expect(role, equals(UserRole.superAdmin));
    });
  });
}
```

### Core Functions:
1. **`void main()`**: Entry point where all test registration happens.
2. **`group(String description, void Function() body)`**: Groups related tests into logical blocks. Can be nested (e.g. `group('UserRole')` -> `group('fromString')`).
3. **`test(String description, void Function() body)`**: An individual test case.
4. **`expect(dynamic actual, dynamic matcher)`**: Verifies that `actual` matches the expectation.

---

## 3. Matchers & Assertions

Common matchers provided by `flutter_test`:

| Matcher | Description | Example |
| :--- | :--- | :--- |
| `equals(expected)` | Checks value equality | `expect(role.value, equals('doctor'));` |
| `isTrue` | Asserts value is `true` | `expect(role.isAdmin, isTrue);` |
| `isFalse` | Asserts value is `false` | `expect(role.isAdmin, isFalse);` |
| `isNull` | Asserts value is `null` | `expect(user.departmentId, isNull);` |
| `isNotNull` | Asserts value is not `null` | `expect(user.id, isNotNull);` |
| `isA<T>()` | Asserts object is of type `T` | `expect(e, isA<FormatException>());` |
| `throwsA(matcher)` | Asserts that a function throws an exception | `expect(() => parse(''), throwsA(isA<FormatException>()));` |
| `contains(element)` | Asserts a collection or string contains element | `expect(['admin', 'doctor'], contains('doctor'));` |

---

## 4. Testing Patterns & Best Practices

### Pattern A: Parameterized / Table-Driven Tests (Dart 3 Records)

Instead of duplicating test functions, define a table of inputs and expected outputs, and loop through them.

> [!TIP]
> Always loop **outside** `test(...)` so that each case registers as an independent test in the test runner.

```dart
group('fromString', () {
  const cases = [
    ('super_admin', UserRole.superAdmin),
    ('department_admin', UserRole.departmentAdmin),
    ('doctor', UserRole.doctor),
    ('viewer', UserRole.viewer),
  ];

  for (final (input, expected) in cases) {
    test('parses "$input" to $expected', () {
      expect(UserRole.fromString(input), equals(expected));
    });
  }
});
```

### Pattern B: Exception / Error Handling (Closures)

When testing that a piece of code throws an exception, **always pass a closure** (`() => functionCall()`) to `expect`.

```dart
test('throws FormatException on unknown role', () {
  expect(
    () => UserRole.fromString('invalid_role'),
    throwsA(isA<FormatException>()),
  );
});
```

*Why a closure?*
If you call `UserRole.fromString('invalid')` directly inside `expect(...)`, Dart executes the method before passing the result to `expect`. The exception crashes the test immediately. Passing a zero-argument function (`() => ...`) allows `expect` to execute it inside an internal `try/catch`.

### Pattern C: Value Equality & `hashCode`

When testing immutable value objects and domain entities, always test:
1. Two distinct instances with the identical fields equal each other (`expect(a, equals(b))`).
2. Their hash codes match (`expect(a.hashCode, equals(b.hashCode))`), which is required for `Set` and `Map` collections.
3. An instance with any differing field is not equal (`expect(a, isNot(equals(c)))`).

```dart
group('equality', () {
  test('identical fields are equal and produce same hashCode', () {
    expect(user1, equals(user2));
    expect(user1.hashCode, equals(user2.hashCode));
  });

  test('different fields are not equal', () {
    expect(user1, isNot(equals(differentUser)));
  });
});
```

### Pattern D: Immutable State Updates (`copyWith`)

Verify that `copyWith` modifies the targeted field while preserving all other properties untouched, and produces an identical copy when called without arguments:

```dart
test('updates specified fields while preserving others', () {
  final updated = original.copyWith(fullName: 'New Name');
  expect(updated.fullName, equals('New Name'));
  expect(updated.id, equals(original.id)); // preserved
});

test('returns identical instance when no arguments passed', () {
  expect(original.copyWith(), equals(original));
});
```

---

## 5. Test-Driven Development (TDD) Lifecycle

1. **RED**: Write the test first before writing production code. Run `flutter test` and verify that the test fails (or fails to compile because the class does not exist yet).
2. **GREEN**: Write the minimal production code needed to make the test pass.
3. **REFACTOR**: Clean up any duplication, improve readability, and ensure clean architecture boundaries without breaking the passing tests.

---

*(This document will be updated as we introduce Mocktail mocks, Riverpod provider testing, and Widget tests.)*
