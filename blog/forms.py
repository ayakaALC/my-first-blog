from django import forms
from .models import Post, Comment



class PostForm(forms.ModelForm):

    class Meta:
        model = Post
        fields = ('title', 'text', 'model_pic')


class CommentForm(forms.ModelForm):

    class Meta:
        model = Comment
        fields = ('author', 'text',)



"""class DocumentForm(forms.ModelForm):
    class Meta:
        model = Document
        fields = ('description', 'document', )"""

"""class HotelForm(forms.ModelForm):

    class Meta:
        model = Post
        fields = ['name', 'hotel_Main_Img']"""

