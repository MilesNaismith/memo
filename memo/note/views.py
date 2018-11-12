from django.shortcuts import render
from .models import Budget, Company
from .processing import get_note
from django.http import FileResponse

# Create your views here.
def note(request):
    if request.method == "POST":
        print(request.POST)
        budget_item = request.POST['budget_items']
        company_name = request.POST['companies']
        pay_reason = request.POST['pay_reason']
        pay_check = request.POST['pay_check']
        pay_sum = request.POST['pay_sum']
        get_note(company_name, budget_item, pay_reason, pay_check, pay_sum)
        response = FileResponse(open('note.docx', 'rb'))
        return response
       # return render(request, 'note/success.html')
    budget_items = Budget.objects.filter()
    companies = Company.objects.filter()
    return render(request, 'note/note.html', {'budget_items': budget_items, 'companies': companies})