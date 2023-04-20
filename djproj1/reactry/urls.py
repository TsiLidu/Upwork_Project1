# this I created it

from django.urls import path
from . import views

urlpatterns = [
         path("", views.index, name = 'index')
]