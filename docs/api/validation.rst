.. _validation_api:

Validation
==========

The validation framework checks a |Presentation| against Open XML schema,
semantic, and package-structure constraints. It is modeled on
``OpenXmlValidator`` from the Microsoft Open XML SDK.

|PresentationValidator| objects
--------------------------------

The validator is constructed directly and used to check an entire
presentation, a single part, or an individual XML element::

    from pptx_ng import Presentation
    from pptx_ng.validation import PresentationValidator

    prs = Presentation('my_deck.pptx')
    validator = PresentationValidator()
    errors = validator.validate(prs)
    for error in errors:
        print(f"[{error.error_type}] {error.description}")

.. autoclass:: pptx.validation.PresentationValidator()
   :members:
   :undoc-members:


|ValidationErrorInfo| objects
------------------------------

Each error returned by validation is a |ValidationErrorInfo| instance.

.. autoclass:: pptx.validation.context.ValidationErrorInfo()
   :members:
   :undoc-members:
