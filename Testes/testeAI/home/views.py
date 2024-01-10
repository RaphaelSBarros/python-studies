from django.shortcuts import render
from .models import Video
from django.http import HttpResponse

# Create your views here.
def home(request):
    if request.method == 'GET':
        return render(request, 'home.html')
    elif request.method == "POST":
        titulo = request.POST.get('titulo')
        video = request.FILES.get('video')
        
        video_upload = Video(titulo=titulo, video=video)
        video_upload.save()
        
        return HttpResponse('Deu Certo')