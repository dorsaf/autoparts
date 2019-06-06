# -*- coding: utf-8 -*-

from odoo import api, exceptions, fields, models, _
import logging
import xlrd
import os
import re
from xlrd import open_workbook
import base64
from odoo.exceptions import UserError, RedirectWarning, ValidationError
_logger = logging.getLogger(__name__)

EXTENSIONS = ['xlsx','xlsm','xltx','xltm','xls']


class ImportVariantWizard(models.TransientModel):
    _name = 'product.import'

class ImportReferenceWizard(models.TransientModel):
    _name = 'reference.import'

    name = fields.Char(string="Nom", required=True, )
    file = fields.Binary(string='Fichier')
    filename = fields.Char(string='Nom de fichier')
    first_row_is_header = fields.Boolean("La première ligne est l'en-tête ?")
    reference_id = fields.Many2many('autopart.reference', string='Reference', readonly=True)

    state = fields.Selection([('draft', 'Brouillon'), ('validate', 'Validé'), ('process', 'Traité')], default='draft')

    def validate_file(self):
        if not self.file or not self.filename:
            raise ValidationError("Vous devez d'abord sélectionner un fichier!")
        extension = self.filename.split('.')[-1]
        if extension not in EXTENSIONS:
            raise ValidationError('Format de fichier non supporté!')

        file_data = base64.b64decode(self.file)

        book = xlrd.open_workbook(file_contents=file_data)
        sheet = book.sheet_by_index(0)

        if not sheet:
            raise ValidationError('Le fichier doit contenir au moins une feuille')

        start_index = 1 if self.first_row_is_header else 0
        if start_index == 1:
            reference = sheet.cell_value(0, 0).lower()
            designation = sheet.cell_value(0, 1).lower()
            pvc = sheet.cell_value(0, 2).lower()
            poids = sheet.cell_value(0, 3).lower()
            ref_fournisseur = sheet.cell_value(0, 4).lower()

            if not reference:
                raise ValidationError("Le nom d'en-tête ne peuvent pas être vides")
            if not designation:
                raise ValidationError("Le nom d'en-tête ne peuvent pas être vides")
            if not pvc:
                raise ValidationError("Le nom d'en-tête ne peuvent pas être vides")
            if not poids:
                raise ValidationError("Le nom d'en-tête ne peuvent pas être vides")
            if not ref_fournisseur:
                raise ValidationError("Le nom d'en-tête ne peuvent pas être vides")
        return sheet



    def _create_references(self, sheet):

        reference_obj = self.env['autopart.reference']

        start_index = 1 if self.first_row_is_header else 0
        if start_index == 1:
            reference = sheet.cell_value(0, 0).lower()
            designation = sheet.cell_value(0, 1).lower()
            pvc = sheet.cell_value(0, 2).lower()
            poids = sheet.cell_value(0, 3).lower()
            ref_fournisseur = sheet.cell_value(0, 4).lower()

            header = [reference,designation,pvc,poids,ref_fournisseur]

            ref_index = header.index('reference')
            designation_index = header.index("designation")
            pvc_index = header.index('pvc')
            poids_index = header.index('poids')
            ref_fournisseur_index = header.index('reference fournisseur')

            for row_number in range(1, sheet.nrows):

                ref_name = sheet.cell_value(row_number, ref_index)
                ref_designation = sheet.cell_value(row_number, designation_index)
                ref_pvc = sheet.cell_value(row_number, pvc_index)
                ref_poids = sheet.cell_value(row_number, poids_index)
                ref_fournisseur = sheet.cell_value(row_number, ref_fournisseur_index)

                # reference data
                reference_id = reference_obj.search([('name', '=', ref_name)])
                _logger.info("____search______reference_id________: %s ", reference_id)
                if not reference_id:
                    vals = {
                        'name': ref_name,
                        'designation': ref_designation,
                        'pvc': ref_pvc
                    }
                    reference_id = reference_obj.sudo().create(vals)
                    _logger.info("____create______reference_id________: %s ", reference_id)
                else:
                    vals = {
                        'name': ref_name,
                        'designation': ref_designation,
                        'pvc': ref_pvc,
                        'poids': ref_poids,
                        'supplier_reference': ref_fournisseur
                    }
                    reference_id.sudo().write(vals)
                    _logger.info("____write______reference_id________: %s ", reference_id)
        if start_index == 0:
            reference = sheet.cell_value(0, 0)
            designation = sheet.cell_value(0, 1)
            pvc = sheet.cell_value(0, 2)
            poids = sheet.cell_value(0, 3)
            ref_fournisseur = sheet.cell_value(0, 4)

            header = [reference, designation, pvc,poids,ref_fournisseur]

            ref_index = header.index(reference)
            designation_index = header.index(designation)
            pvc_index = header.index(pvc)
            poids_index = header.index(poids)
            ref_fournisseur_index = header.index(ref_fournisseur)

            for row_number in range(0, sheet.nrows):
                ref_name = sheet.cell_value(row_number, ref_index)
                ref_designation = sheet.cell_value(row_number, designation_index)
                ref_pvc = sheet.cell_value(row_number, pvc_index)
                ref_poids = sheet.cell_value(row_number, poids_index)
                fournisseur = sheet.cell_value(row_number, ref_fournisseur_index)
                # reference data
                reference_id = reference_obj.search([('name', '=', ref_name)])
                _logger.info("____search______reference_id________: %s ", reference_id)
                if not reference_id:
                    vals = {
                        'name': ref_name,
                        'designation': ref_designation,
                        'pvc': ref_pvc,
                        'poids': ref_poids,
                        'supplier_reference': fournisseur
                    }
                    reference_id = reference_obj.sudo().create(vals)
                    _logger.info("____create______reference_id________: %s ", reference_id)
                else:
                    vals = {
                        'name': ref_name,
                        'designation': ref_designation,
                        'pvc': ref_pvc,
                        'poids': ref_poids,
                        'supplier_reference': fournisseur
                    }
                    reference_id.sudo().write(vals)
                    _logger.info("____write______reference_id________: %s ", reference_id)



    def validate_files(self):
        sheet = self.validate_file()
        self.state = 'validate'

    def process_file(self):
        sheet = self.validate_file()
        self._create_references(sheet)
        self.state = 'process'
        return True